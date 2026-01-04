fn main() {
    // Compile Windows resources (icon, manifest, toolbar bitmap)
    embed_resource::compile("resources/simpliview.rc", embed_resource::NONE);

    // Link required Windows libraries
    println!("cargo:rustc-link-lib=comctl32");
    println!("cargo:rustc-link-lib=d2d1");
    println!("cargo:rustc-link-lib=dxguid");
    println!("cargo:rustc-link-lib=windowscodecs");
    println!("cargo:rustc-link-lib=ole32");
    println!("cargo:rustc-link-lib=shell32");
    println!("cargo:rustc-link-lib=uxtheme");

    // Rerun if resources change
    println!("cargo:rerun-if-changed=resources/simpliview.rc");
    println!("cargo:rerun-if-changed=resources/simpliview.manifest");
    println!("cargo:rerun-if-changed=photo_portrait.ico");
    println!("cargo:rerun-if-changed=skn16g.png");
}
