//! MSIX conformance: Windows' parser ("Error in parsing app package") rejects
//! packages whose AppxBlockMap.xml File Names use forward slashes --
//! MakeAppx writes BACKSLASH separators (verified 223/223 nested entries
//! against a Microsoft-built Windows Terminal .msix, 0 forward-slashed),
//! while the zip entries themselves stay forward-slashed. This test encodes
//! that rule plus the rest of the structural contract (64 KiB sha256 blocks,
//! LfhSize agreement with the zip local headers, required manifest elements,
//! PE machine type matching ProcessorArchitecture) so a regression fails
//! here instead of on a user's Windows box.
//!
//! Two tiers: (1) always runs -- packages a dummy exe through the real
//! scripts/package-msix.py and validates the output (skips loudly if
//! python3/PIL is unavailable); (2) validates dist/corro-*.msix when present
//! (skips loudly when the release artifacts haven't been built).

use std::collections::HashMap;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::Command;

// ---------------------------------------------------------------- sha-256
// (inline so the test stays std+zip only)

const SHA_K: [u32; 64] = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1,
    0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3,
    0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786,
    0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147,
    0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13,
    0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b,
    0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a,
    0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
    0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
];

fn sha256(data: &[u8]) -> [u8; 32] {
    let mut h = [
        0x6a09e667u32, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
        0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
    ];
    let mut msg = data.to_vec();
    let bitlen = (data.len() as u64).wrapping_mul(8);
    msg.push(0x80);
    while msg.len() % 64 != 56 {
        msg.push(0);
    }
    msg.extend_from_slice(&bitlen.to_be_bytes());
    for chunk in msg.chunks_exact(64) {
        let mut w = [0u32; 64];
        for i in 0..16 {
            w[i] = u32::from_be_bytes([
                chunk[4 * i], chunk[4 * i + 1], chunk[4 * i + 2], chunk[4 * i + 3],
            ]);
        }
        for i in 16..64 {
            let s0 = w[i - 15].rotate_right(7) ^ w[i - 15].rotate_right(18) ^ (w[i - 15] >> 3);
            let s1 = w[i - 2].rotate_right(17) ^ w[i - 2].rotate_right(19) ^ (w[i - 2] >> 10);
            w[i] = w[i - 16]
                .wrapping_add(s0)
                .wrapping_add(w[i - 7])
                .wrapping_add(s1);
        }
        let (mut a, mut b, mut c, mut d, mut e, mut f, mut g, mut hh) =
            (h[0], h[1], h[2], h[3], h[4], h[5], h[6], h[7]);
        for i in 0..64 {
            let s1 = e.rotate_right(6) ^ e.rotate_right(11) ^ e.rotate_right(25);
            let ch = (e & f) ^ ((!e) & g);
            let t1 = hh
                .wrapping_add(s1)
                .wrapping_add(ch)
                .wrapping_add(SHA_K[i])
                .wrapping_add(w[i]);
            let s0 = a.rotate_right(2) ^ a.rotate_right(13) ^ a.rotate_right(22);
            let maj = (a & b) ^ (a & c) ^ (b & c);
            let t2 = s0.wrapping_add(maj);
            hh = g;
            g = f;
            f = e;
            e = d.wrapping_add(t1);
            d = c;
            c = b;
            b = a;
            a = t1.wrapping_add(t2);
        }
        h[0] = h[0].wrapping_add(a);
        h[1] = h[1].wrapping_add(b);
        h[2] = h[2].wrapping_add(c);
        h[3] = h[3].wrapping_add(d);
        h[4] = h[4].wrapping_add(e);
        h[5] = h[5].wrapping_add(f);
        h[6] = h[6].wrapping_add(g);
        h[7] = h[7].wrapping_add(hh);
    }
    let mut out = [0u8; 32];
    for (i, v) in h.iter().enumerate() {
        out[4 * i..4 * i + 4].copy_from_slice(&v.to_be_bytes());
    }
    out
}

fn b64_encode(data: &[u8]) -> String {
    const T: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut s = String::new();
    for c in data.chunks(3) {
        let n = (c[0] as u32) << 16 | (*c.get(1).unwrap_or(&0) as u32) << 8 | (*c.get(2).unwrap_or(&0) as u32);
        s.push(T[((n >> 18) & 63) as usize] as char);
        s.push(T[((n >> 12) & 63) as usize] as char);
        s.push(if c.len() > 1 { T[((n >> 6) & 63) as usize] as char } else { '=' });
        s.push(if c.len() > 2 { T[(n & 63) as usize] as char } else { '=' });
    }
    s
}

// ------------------------------------------------------------ zip walking

fn u16le(b: &[u8], off: usize) -> u16 {
    u16::from_le_bytes([b[off], b[off + 1]])
}

fn u32le(b: &[u8], off: usize) -> u32 {
    u32::from_le_bytes([b[off], b[off + 1], b[off + 2], b[off + 3]])
}

/// Walk zip local headers: name -> local-header size (30 + fn + extra).
fn local_headers(raw: &[u8]) -> HashMap<String, usize> {
    let mut map = HashMap::new();
    let mut pos = 0usize;
    while pos + 30 <= raw.len() {
        if &raw[pos..pos + 4] != b"PK\x03\x04" {
            break;
        }
        let fnlen = u16le(raw, pos + 26) as usize;
        let exlen = u16le(raw, pos + 28) as usize;
        let comp = u32le(raw, pos + 18) as usize;
        let name = String::from_utf8_lossy(&raw[pos + 30..pos + 30 + fnlen]).into_owned();
        map.insert(name, 30 + fnlen + exlen);
        pos += 30 + fnlen + exlen + comp;
    }
    map
}

fn attr(tag: &str, key: &str) -> String {
    let pat = format!("{key}=\"");
    let s = tag.find(&pat).unwrap_or_else(|| panic!("attr {key} missing in {tag}"));
    let rest = &tag[s + pat.len()..];
    rest[..rest.find('"').expect("unterminated attr")].to_string()
}

struct BlockFile {
    name: String,
    size: usize,
    lfh: usize,
    hashes: Vec<String>,
}

fn parse_blockmap(xml: &str) -> Vec<BlockFile> {
    let mut out = Vec::new();
    for part in xml.split("<File ").skip(1) {
        let end = part.find("</File>").expect("unclosed File element");
        let body = &part[..end];
        let head = &body[..body.find('>').expect("unclosed File tag")];
        let hashes: Vec<String> = body
            .split("<Block Hash=\"")
            .skip(1)
            .map(|b| b[..b.find('"').expect("unterminated hash")].to_string())
            .collect();
        out.push(BlockFile {
            name: attr(head, "Name"),
            size: attr(head, "Size").parse().expect("bad Size"),
            lfh: attr(head, "LfhSize").parse().expect("bad LfhSize"),
            hashes,
        });
    }
    out
}

// ------------------------------------------------------------- validation

fn zip_contents(path: &Path) -> HashMap<String, Vec<u8>> {
    let f = std::fs::File::open(path).unwrap();
    let mut za = zip::ZipArchive::new(f).unwrap();
    let mut map = HashMap::new();
    for i in 0..za.len() {
        let mut e = za.by_index(i).unwrap();
        let mut v = Vec::new();
        e.read_to_end(&mut v).unwrap();
        map.insert(e.name().to_string(), v);
    }
    map
}

fn pe_machine(exe: &[u8], what: &str) -> u16 {
    assert!(exe.len() > 64 && &exe[0..2] == b"MZ", "{what}: not a PE (no MZ)");
    let e_lfanew = u32le(exe, 0x3c) as usize;
    assert!(e_lfanew + 6 <= exe.len(), "{what}: truncated PE header");
    assert!(&exe[e_lfanew..e_lfanew + 4] == b"PE\0\0", "{what}: no PE signature");
    u16le(exe, e_lfanew + 4)
}

/// Full structural validation of one .msix. Panics with context on failure.
fn validate_msix(path: &Path, expect_arch: Option<&str>, check_pe: bool) {
    let what = path.display().to_string();
    let raw = std::fs::read(path).unwrap();
    let entries = zip_contents(path);
    for root in ["[Content_Types].xml", "AppxManifest.xml", "AppxBlockMap.xml"] {
        assert!(entries.contains_key(root), "{what}: missing {root}");
    }
    let manifest = String::from_utf8(entries["AppxManifest.xml"].clone()).unwrap();
    let arch = attr(&manifest[manifest.find("<Identity").unwrap()..], "ProcessorArchitecture");
    if let Some(want) = expect_arch {
        assert_eq!(arch, want, "{what}: manifest arch mismatch");
    }
    for needle in [
        "<DisplayName>",
        "<PublisherDisplayName>",
        "<Logo>",
        "<Resource Language=",
        "TargetDeviceFamily",
        "runFullTrust",
        "EntryPoint=\"Windows.FullTrustApplication\"",
        "Square150x150Logo=",
        "Square44x44Logo=",
    ] {
        assert!(manifest.contains(needle), "{what}: manifest missing {needle}");
    }
    let exe_tag = &manifest[manifest.find("Executable=\"").unwrap()..];
    let exe_name = attr(exe_tag, "Executable");
    assert!(entries.contains_key(&exe_name), "{what}: manifest exe {exe_name} not in package");

    // blockmap rules: backslash separators, sizes, 64 KiB sha256 blocks, LfhSize
    let blockmap = String::from_utf8(entries["AppxBlockMap.xml"].clone()).unwrap();
    assert!(
        blockmap.contains("http://www.w3.org/2001/04/xmlenc#sha256"),
        "{what}: blockmap must be sha256"
    );
    let lfh = local_headers(&raw);
    let mut files = 0usize;
    for bf in parse_blockmap(&blockmap) {
        assert!(
            !bf.name.contains('/'),
            "{what}: blockmap File Name must use backslashes (Windows rejects forward slashes): {}",
            bf.name
        );
        let zname = bf.name.replace('\\', "/");
        let data = entries.get(&zname).unwrap_or_else(|| panic!("{what}: blockmap {zname} not in zip"));
        assert_eq!(data.len(), bf.size, "{what}: size mismatch for {}", bf.name);
        let expect_blocks = if data.is_empty() { 0 } else { (data.len() + 65535) / 65536 };
        assert_eq!(bf.hashes.len(), expect_blocks, "{what}: 64 KiB block count wrong for {}", bf.name);
        for (i, h) in bf.hashes.iter().enumerate() {
            let end = ((i + 1) * 65536).min(data.len());
            assert_eq!(&b64_encode(&sha256(&data[i * 65536..end])), h, "{what}: block {i} hash wrong for {}", bf.name);
        }
        let got_lfh = lfh.get(&zname).unwrap_or_else(|| panic!("{what}: no local header for {zname}"));
        assert_eq!(*got_lfh, bf.lfh, "{what}: LfhSize wrong for {}", bf.name);
        files += 1;
    }
    assert!(files >= 3, "{what}: suspiciously few blockmap files ({files})");

    if check_pe {
        let machine = pe_machine(&entries[&exe_name], &what);
        let want_machine = match arch.as_str() {
            "x86" => 0x014c,
            "x64" => 0x8664,
            "arm64" => 0xaa64,
            _ => panic!("{what}: unknown arch {arch}"),
        };
        assert_eq!(machine, want_machine, "{what}: PE machine type does not match {arch}");
    }
}

fn repo_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).to_path_buf()
}

fn have_python_packager() -> bool {
    Command::new("python3")
        .args(["-c", "import PIL, zipfile, hashlib"])
        .output()
        .map(|o| o.status.success())
        .unwrap_or(false)
}

#[test]
fn msix_dummy_package_conforms() {
    if !have_python_packager() {
        eprintln!("SKIP msix_dummy_package_conforms: python3+PIL unavailable");
        return;
    }
    let dir = std::env::temp_dir().join(format!("corro-msix-test-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    for arch in ["x86", "x64", "arm64"] {
        std::fs::write(dir.join(format!("fake-{arch}.exe")), b"MZ-fake-exe-payload").unwrap();
    }
    let root = repo_root();
    let st = Command::new("python3")
        .arg(root.join("scripts/package-msix.py"))
        .arg("--exe-x86")
        .arg(dir.join("fake-x86.exe"))
        .arg("--exe-x64")
        .arg(dir.join("fake-x64.exe"))
        .arg("--exe-arm64")
        .arg(dir.join("fake-arm64.exe"))
        .arg("--outdir")
        .arg(&dir)
        .arg("--publisher")
        .arg("CN=Test")
        .arg("--identity")
        .arg("Test.Corro")
        .current_dir(&root)
        .status()
        .unwrap();
    assert!(st.success(), "package-msix.py failed on dummy exes");
    for arch in ["x86", "x64", "arm64"] {
        let matches: Vec<PathBuf> = std::fs::read_dir(&dir)
            .unwrap()
            .filter_map(|e| {
                let p = e.unwrap().path();
                (p.extension().map(|x| x == "msix").unwrap_or(false)
                    && p.file_name().unwrap().to_string_lossy().ends_with(&format!("-{arch}.msix")))
                .then_some(p)
            })
            .collect();
        assert_eq!(matches.len(), 1, "expected one -{arch}.msix in {dir:?}");
        validate_msix(&matches[0], Some(arch), false);
    }
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn msix_dist_packages_conform() {
    let dist = repo_root().join("dist");
    let mut found = 0usize;
    if let Ok(rd) = std::fs::read_dir(&dist) {
        for e in rd.flatten() {
            let p = e.path();
            let name = p.file_name().unwrap().to_string_lossy().into_owned();
            if !name.ends_with(".msix") {
                continue;
            }
            let arch = name.trim_end_matches(".msix").rsplit('-').next().unwrap().to_string();
            assert!(
                ["x86", "x64", "arm64"].contains(&arch.as_str()),
                "unexpected msix arch in {name}"
            );
            validate_msix(&p, Some(&arch), true);
            found += 1;
        }
    }
    if found == 0 {
        eprintln!(
            "SKIP msix_dist_packages_conform: no dist/*.msix present; build them with scripts/package-msix.py"
        );
    } else {
        // Require every present package to be valid (done above), but do not
        // demand all three arches: arm64 needs an aarch64-Windows unwinder
        // that is not available in every build environment, so a release
        // built without one legitimately ships x86+x64 only. Warn so the gap
        // is visible rather than failing a package that is itself correct.
        let archs: std::collections::BTreeSet<String> = std::fs::read_dir(&dist)
            .unwrap()
            .flatten()
            .filter_map(|e| {
                let n = e.file_name().to_string_lossy().into_owned();
                n.ends_with(".msix")
                    .then(|| n.trim_end_matches(".msix").rsplit('-').next().unwrap().to_string())
            })
            .collect();
        eprintln!("validated {found} msix package(s): {archs:?}");
        if !archs.contains("arm64") {
            eprintln!(
                "NOTE: no arm64 .msix in dist/ — the aarch64-pc-windows-gnullvm link \
                 needs an aarch64 Windows unwinder (see docs/binaries.txt)"
            );
        }
    }
}
