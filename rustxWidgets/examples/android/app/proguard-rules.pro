# Keep the JNI entry points the Kotlin/Java side declares. R8 cannot see
# them from Rust, so they must be named here or a release build will strip
# them and crash with UnsatisfiedLinkError.
-keepclasseswithmembernames,includedescriptorclasses class * {
    native <methods>;
}

# Keep the app's own shim classes: they are instantiated through JNI
# (ClassLoader.loadClass + new_object) rather than referenced from Java, so
# R8 has no call site to reach them from.
-keep class com.example.MainActivity { *; }
-keep class com.example.RustCallback { *; }
