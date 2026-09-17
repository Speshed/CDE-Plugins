from pathlib import Path

from tools.verify_single_exe import verify_distribution

GOOD_TOC = """(([('PySide6\\\\QtWebEngineProcess.exe','x','BINARY')], [('PySide6\\\\resources\\\\icudtl.dat','x','DATA'), ('PySide6\\\\resources\\\\qtwebengine_resources.pak','x','DATA'), ('PySide6\\\\translations\\\\qtwebengine_locales\\\\ru.pak','x','DATA')]))"""


def fixture(tmp_path: Path, toc: str = GOOD_TOC):
    release = tmp_path / "release"
    build = tmp_path / "build" / "portable-work" / "Larix_CDE"
    release.mkdir(parents=True)
    build.mkdir(parents=True)
    (release / "Larix_CDE.exe").write_bytes(b"x")
    (build / "PKG-00.toc").write_text(toc, encoding="utf-8")
    return release, tmp_path / "build" / "portable-work"


def test_valid_fixture(tmp_path: Path):
    release, build = fixture(tmp_path)
    assert verify_distribution(release, build) == []


def test_rejects_invalid_structure_and_resources(tmp_path: Path):
    toc = GOOD_TOC.replace("QtWebEngineProcess.exe", "VitroCAD.exe").replace("ru.pak", "en-US.pak").replace("icudtl.dat", "v8_context_snapshot.debug.bin")
    release, build = fixture(tmp_path, toc)
    (release / "assets").mkdir()
    errors = verify_distribution(release, build)
    assert any("только Larix_CDE.exe" in error for error in errors)
    assert any("вложенные EXE" in error for error in errors)
    assert any("только qtwebengine" in error for error in errors)
    assert any("debug-ресурсы" in error for error in errors)


def test_rejects_size_limit(tmp_path: Path):
    release, build = fixture(tmp_path)
    assert any("превышает лимит" in error for error in verify_distribution(release, build, max_bytes=0))
