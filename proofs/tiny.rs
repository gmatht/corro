use vstd::prelude::*;
verus! {
spec fn pow26(n: nat) -> int decreases n { if n <= 0 { 1 } else { 26 * pow26((n - 1) as nat) } }
proof fn t(k: nat) ensures pow26(k) >= 1, decreases k { if k > 0 { t((k-1) as nat); } }
} // verus
fn main() {}
