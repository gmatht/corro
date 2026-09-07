//! C99 float shims for the vintage MSVC toolchain (rust9x targets only).
//!
//! On MSVC targets, `compiler_builtins` and rustc lower float intrinsics
//! (`f64::round`, `f64::trunc`, `f64::cbrt`, `f64::log2`, `f64::mul_add`, …)
//! to CRT libcalls. The VC6 (MSVC 6.0) static CRT we link for the rust9x
//! targets is the most Win9x-compatible, but it predates C99 — so the
//! symbols `round`, `trunc`, `cbrt`, `log2`, `fma` (and f32 variants) end up
//! undefined at link time. This module provides them.
//!
//! Modelled on /root/src/lsl-usb/rust9x/nwg-test/src/float_shim.rs (which
//! only needed `round`/`roundf`); corro also uses trunc/cbrt/log2/fma.
//!
//! IMPORTANT: none of these implementations may call an f64/f32 intrinsic
//! that itself lowers to a C99 CRT libcall (`round()`, `trunc()`, `fma()`,
//! `mul_add()`, …) or the link would recurse. Only bit manipulation, plain
//! arithmetic and saturating casts are used.
//!
//! Only compiled for rust9x msvc targets (see lib.rs); every binary links
//! the corro rlib, so the linker pulls these objects in to satisfy the
//! undefined symbols.

// ---------- f64: trunc (bit manipulation, exact) ----------

#[no_mangle]
pub extern "C" fn trunc(x: f64) -> f64 {
    let b = x.to_bits();
    let biased = ((b >> 52) & 0x7ff) as i32;
    let frac = b & 0x000f_ffff_ffff_ffff;
    if biased == 0x7ff || (biased == 0 && frac == 0) {
        return x; // inf / nan / ±0
    }
    let e = biased - 1023;
    if e >= 52 {
        return x; // already integral
    }
    if e < 0 {
        return if x < 0.0 { -0.0 } else { 0.0 }; // |x| < 1
    }
    let mask = (1u64 << (52 - e)) - 1;
    f64::from_bits(b & !mask)
}

// ---------- f64: round (half away from zero, like C99 round) ----------

const F64_BIG: f64 = 9_007_199_254_740_992.0; // 2^53 — |x| >= this is integral

#[no_mangle]
pub extern "C" fn round(x: f64) -> f64 {
    if x.is_nan() || x.abs() >= F64_BIG {
        return x;
    }
    let t = trunc(x);
    let frac = x - t; // exact (Sterbenz) for the ranges involved
    if x >= 0.0 {
        if frac >= 0.5 {
            t + 1.0
        } else {
            t
        }
    } else if frac <= -0.5 {
        t - 1.0
    } else {
        t
    }
}

// ---------- f64: cbrt (Newton iteration, no libm) ----------

const TWO_POW_52: f64 = 4_503_599_627_370_496.0; // 2^52
const F64_MIN_NORMAL: f64 = 2.2250738585072014e-308; // 2^-1022

#[no_mangle]
pub extern "C" fn cbrt(x: f64) -> f64 {
    if x == 0.0 || x.is_nan() || x.is_infinite() {
        return x;
    }
    let neg = x < 0.0;
    let mut ax = x.abs();
    // lift subnormals into the normal range (compensated at the end)
    let mut scale = 0i32;
    if ax < F64_MIN_NORMAL {
        ax *= TWO_POW_52;
        scale -= 52;
    }
    // decompose ax = m2 * 2^(3q), with m2 in [1,8)
    let bits = ax.to_bits();
    let e = (((bits >> 52) & 0x7ff) as i32) - 1022; // ax = m * 2^e, m in [1,2)
    let m = f64::from_bits((bits & !(0x7ffu64 << 52)) | (1023u64 << 52));
    let q = e.div_euclid(3);
    let r = e.rem_euclid(3);
    let m2 = m * (1u64 << r) as f64; // in [1, 8)
    // want y with y^3 = m2, y in [1,2); start from a linear approx, then
    // Newton: y' = (2y + m2/y^2) / 3 (converges globally for positive y)
    let mut y = 1.0 + (m2 - 1.0) / 3.0;
    for _ in 0..12 {
        y = (2.0 * y + m2 / (y * y)) / 3.0;
    }
    let exp2q = f64::from_bits(((q + 1023) as u64) << 52);
    let res = y * exp2q * f64::from_bits(((scale + 1023) as u64) << 52);
    if neg {
        -res
    } else {
        res
    }
}

// ---------- f64: log2 (mantissa decomposition + atanh series) ----------

const LN2_INV: f64 = 1.4426950408889634; // 1 / ln(2)

#[no_mangle]
pub extern "C" fn log2(x: f64) -> f64 {
    if x.is_nan() {
        return x;
    }
    if x < 0.0 {
        return f64::NAN;
    }
    if x == 0.0 {
        return f64::NEG_INFINITY;
    }
    if x.is_infinite() {
        return x;
    }
    let mut m = x;
    let mut k = 0i32;
    if m < F64_MIN_NORMAL {
        m *= TWO_POW_52; // lift subnormals
        k -= 52;
    }
    // m = mant * 2^e with mant in [1,2)
    let bits = m.to_bits();
    let e = (((bits >> 52) & 0x7ff) as i32) - 1023;
    let mant = f64::from_bits((bits & !(0x7ffu64 << 52)) | (1023u64 << 52));
    // log2(mant) = 2/ln2 * atanh((mant-1)/(mant+1)); z <= 1/3 so the series
    // converges fast (each term shrinks ~9x)
    let z = (mant - 1.0) / (mant + 1.0);
    let z2 = z * z;
    let mut sum = 0.0;
    let mut p = z;
    let mut n = 1.0f64;
    for _ in 0..30 {
        sum += p / n;
        p *= z2;
        n += 2.0;
    }
    e as f64 + k as f64 + 2.0 * sum * LN2_INV
}

// ---------- f64: fma (Dekker 2-product + 2-sum, no fma intrinsic) ----------

#[no_mangle]
pub extern "C" fn fma(a: f64, b: f64, c: f64) -> f64 {
    let p = a * b;
    // Special cases: anything non-finite, or a degenerate product, is handled
    // exactly (well-defined) by a plain p + c. Also the extreme-|a|/|b| range
    // where Dekker's split would overflow: p + c differs from a true fma only
    // in cases where the correction term is below the rounding threshold
    // anyway (and p itself would be inf/nan there).
    if !p.is_finite() || !c.is_finite() || a == 0.0 || b == 0.0 || a.abs() > 1e300 || b.abs() > 1e300
    {
        return p + c;
    }
    // exact product p + err via Dekker's split (2^27+1)
    const SPLIT: f64 = 134_217_729.0;
    let ca = SPLIT * a;
    let ahi = ca - (ca - a);
    let alo = a - ahi;
    let cb = SPLIT * b;
    let bhi = cb - (cb - b);
    let blo = b - bhi;
    let err = ((ahi * bhi - p) + ahi * blo + alo * bhi) + alo * blo;
    // exact sum s + err2 of p and c via two-sum
    let s = p + c;
    let bv = s - p;
    let err2 = (p - (s - bv)) + (c - bv);
    s + (err + err2)
}

// ---------- f32 variants ----------
// truncf/roundf implemented directly (exact); the transcendentals go through
// the f64 versions (double rounding can cost at most ~1 f32 ulp, fine here).

const F32_BIG: f32 = 8_388_608.0; // 2^23 — |x| >= this is integral

#[no_mangle]
pub extern "C" fn truncf(x: f32) -> f32 {
    let b = x.to_bits();
    let biased = ((b >> 23) & 0xff) as i32;
    let frac = b & 0x007f_ffff;
    if biased == 0xff || (biased == 0 && frac == 0) {
        return x;
    }
    let e = biased - 127;
    if e >= 23 {
        return x;
    }
    if e < 0 {
        return if x < 0.0 { -0.0 } else { 0.0 };
    }
    let mask = (1u32 << (23 - e)) - 1;
    f32::from_bits(b & !mask)
}

#[no_mangle]
pub extern "C" fn roundf(x: f32) -> f32 {
    if x.is_nan() || x.abs() >= F32_BIG {
        return x;
    }
    let t = truncf(x);
    let frac = x - t;
    if x >= 0.0 {
        if frac >= 0.5 {
            t + 1.0
        } else {
            t
        }
    } else if frac <= -0.5 {
        t - 1.0
    } else {
        t
    }
}

#[no_mangle]
pub extern "C" fn cbrtf(x: f32) -> f32 {
    cbrt(x as f64) as f32
}

#[no_mangle]
pub extern "C" fn log2f(x: f32) -> f32 {
    log2(x as f64) as f32
}

#[no_mangle]
pub extern "C" fn fmaf(a: f32, b: f32, c: f32) -> f32 {
    fma(a as f64, b as f64, c as f64) as f32
}
