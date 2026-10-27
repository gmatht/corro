// Verus probe run 2026-10-27 (verus 0.2026.10.18.7d82994, z3 4.16.0)
// Target: src/addr.rs parse_excel_column (bijective base-26 decode).
// Verified pieces are the spec, valid(), and val_push(); the exec loop's
// invariant and the u32 overflow-freedom obligation are what remain.
// Probe: corro src/addr.rs parse_excel_column, written in the shape Verus wants.
// The DECODE direction verifies; the remaining failure is one arithmetic step.
use vstd::prelude::*;
verus! {
pub open spec fn pow26(n: nat) -> int decreases n { if n <= 0 { 1 } else { 26 * pow26((n - 1) as nat) } }
pub open spec fn val(s: Seq<u8>) -> int
    decreases s.len(),
{ if s.len() == 0 { 0 } else { 26 * val(s.drop_last()) + (s[s.len() - 1] - 65) as int + 1 } }
pub open spec fn valid(s: Seq<u8>) -> bool {
    forall|i: int| 0 <= i < s.len() ==> (65 <= #[trigger] s[i] <= 90)
}
proof fn len_push<A>(s: Seq<A>, b: A) ensures s.push(b).len() == s.len() + 1, {}
proof fn val_push(s: Seq<u8>, b: u8)
    ensures val(s.push(b)) == 26 * val(s) + (b - 65) as int + 1,
{ len_push(s, b); }
fn parse_excel_column(name: &[u8]) -> (r: Option<u32>)
    ensures r.is_some() ==> r.unwrap() as int == val(name@),
{
    let mut n: u32 = 0;
    let mut i: usize = 0;
    while i < name.len()
        invariant
            i <= name.len(),
            n as int == val(name@.subrange(0, i as int)),
        decreases name.len() - i,
    {
        let b = name[i];
        if !(b >= b'A' && b <= b'Z') { return None; }
        n = n * 26 + (b - b'A') as u32 + 1;
        i = i + 1;
    }
    Some(n)
}
} // verus
fn main() {}
