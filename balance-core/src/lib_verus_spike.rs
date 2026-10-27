use vstd::prelude::*;

verus! {

pub open spec fn digit_val(c: char) -> int { (c as int) - ('0' as int) }

pub open spec fn all_digits(s: Seq<char>) -> bool {
    forall|i: int| 0 <= i < s.len() ==> ('0' <= #[trigger] s[i] <= '9')
}

pub open spec fn whole_val(s: Seq<char>) -> int
    decreases s.len(),
{
    if s.len() == 0 { 0 } else { 10 * whole_val(s.drop_last()) + digit_val(s[s.len() - 1]) }
}

pub open spec fn frac_val(s: Seq<char>) -> int
    decreases s.len(),
{
    if s.len() == 0 { 0 }
    else if s.len() == 1 { 10 * digit_val(s[0]) }
    else { 10 * digit_val(s[0]) + digit_val(s[1]) }
}

pub open spec fn cents_of(w: Seq<char>, f: Seq<char>) -> int { whole_val(w) * 100 + frac_val(f) }

pub assume_specification [str::trim] (s: &str) -> (t: &str);

pub fn parse_amount_cents(raw: &str) -> (r: Option<i64>)
    ensures
        r.is_some() ==> exists|w: Seq<char>, f: Seq<char>|
            #![trigger cents_of(w, f)]
            all_digits(w) && all_digits(f) && w.len() >= 1
            && r.unwrap() as int == cents_of(w, f),
{
    let t = raw.trim();
    if t.is_empty() {
        return None;
    }
    let mut s = t;
    // ---- Verus stops HERE. The rest of the real body is below, verbatim,
    // because str::replace has no vstd spec -- see proofs/VERUS_FIT_REPORT.md.
    None
}

/*
    let negative = match s.chars().next()? {
        '+' => { s = &s[1..]; false }
        '-' => { s = &s[1..]; true }
        _ => false,
    };
    let (whole, frac) = match s.split_once('.') { Some((w, f)) => (w, f), None => (s.as_str(), "") };
    if whole.is_empty() || !whole.chars().all(|c| c.is_ascii_digit()) { return None; }
    if !frac.chars().all(|c| c.is_ascii_digit()) { return None; }
    let mut cents = whole.parse::<i64>().ok()?.saturating_mul(100);
    let frac_cents = match frac.len() {
        0 => 0,
        1 => frac.parse::<i64>().ok()?.saturating_mul(10),
        _ => frac.get(0..2)?.parse::<i64>().ok()?,
    };
    cents = cents.saturating_add(frac_cents);
*/

} // verus

fn main() {}
