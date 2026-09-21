plugins {
    id("com.android.application")
    kotlin("android") version "1.9.0" apply false
}

android {
    namespace = "com.example"
    // Play requires new apps and updates to target API 36 (Android 16) as of
    // 2026-08-31; see rustxWidgets/docs/ANDROID_GUIDELINES.md.
    compileSdk = 36

    defaultConfig {
        applicationId = "com.example"
        minSdk = 24
        targetSdk = 36
        versionCode = 1
        versionName = "1.0"

        ndk {
            // 64-bit is mandatory for Play; keep armeabi-v7a only if you need
            // pre-2015 devices. See ANDROID_GUIDELINES.md §"64-bit".
            abiFilters += listOf("arm64-v8a", "x86_64")
        }
    }

    buildTypes {
        release {
            // R8: shrink + obfuscate, and keep the mapping file to upload to
            // Play so production stack traces stay readable.
            isMinifyEnabled = true
            isShrinkResources = true
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro",
            )
        }
    }

    sourceSets {
        getByName("main") {
            jniLibs.srcDirs(
                "../../../rswidgets/target/arm64-v8a/release/",
                "../../../rswidgets/target/x86_64/release/",
            )
        }
    }
}

dependencies {
    // Material 3 components: the generated theme parents against
    // Theme.Material3.*, which lives in this artifact. rswidgets does not
    // fetch it for you (see ANDROID_GUIDELINES.md §1).
    implementation("com.google.android.material:material:1.12.0")
    implementation("androidx.appcompat:appcompat:1.6.1")
}
