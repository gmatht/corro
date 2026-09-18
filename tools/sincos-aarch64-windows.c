// mingwex-compatible sincos() shim for the zig-linked Windows/ARM64 build.
// num-complex emits a `sincos` reference (GNU extension); real MinGW-w64
// provides it in libmingwex, which zig's bundled CRT does not ship. sin/cos
// resolve to the real UCRT at link time; only the joint entry point is new.
void sincos(double x, double *s, double *c) {
    *s = __builtin_sin(x);
    *c = __builtin_cos(x);
}
