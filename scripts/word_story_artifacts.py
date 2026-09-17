"""Static artifact binding shared by story native comparisons."""
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def build_hash():
    folder = ROOT / 'build'
    files = sorted(p.relative_to(folder).as_posix() for p in folder.rglob('*') if p.is_file())
    assert files, 'Missing static build'
    manifest = ''.join(f'{name}\0{hashlib.sha256((folder / name).read_bytes()).hexdigest()}\n' for name in files)
    return hashlib.sha256(manifest.encode('utf-8')).hexdigest()


def check_build(report):
    assert report.get('buildHash') == build_hash(), 'Browser evidence belongs to a different static build'
