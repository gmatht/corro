# pdcurses-sys [![Build status](https://ci.appveyor.com/api/projects/status/7quldtl11lsitu2v?svg=true)](https://ci.appveyor.com/project/ihalila/pdcurses-sys) [![Crates.io](https://img.shields.io/crates/v/pdcurses-sys.svg)](https://crates.io/crates/pdcurses-sys)

pdcurses-sys provides Rust FFI bindings for [PDCurses](http://wmcbrine.com/pdcurses/),
specifically the fork by [Bill-Gray](https://github.com/Bill-Gray/PDCurses).

## Requirements

A native C compiler that [cc-rs](https://github.com/alexcrichton/cc-rs)
can use to compile PDCurses.

On Windows this means that you need the Visual C++ Build Tools. Check the [rustup docs](https://github.com/rust-lang-nursery/rustup.rs/blob/master/README.md#working-with-rust-on-windows)
for more Rust <-> Windows information.

## Usage

Cargo.toml
```toml
[dependencies]
pdcurses-sys = "0.7"
```

## Picking a flavor

pdcurses-sys supports both the win32 and win32a flavors of PDCurses. It defaults to win32a if none
is specified. You can use Cargo features to choose which one to use:

Cargo.toml
```toml
[dependencies.pdcurses-sys]
version = "0.7"
features = ["win32"]
```

## License

Licensed under the MIT license, see [LICENSE.md](LICENSE.md)
