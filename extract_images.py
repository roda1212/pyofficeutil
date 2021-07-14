import argparse
import glob
import os
import lib.Log as Log
import zipfile
import pathlib

def InitArgParser() -> argparse.ArgumentParser:
    """
    引数の初期化
    """
    parser = argparse.ArgumentParser(description='OfficeファイルのGrep検索')
    parser.add_argument('target', type=str, help='対象ディレクトリ')
    parser.add_argument('-outdir', type=str, help='結果ディレクトリ', default='result_images')
    return parser

def EnumOfficeFiles(target):
    xlsm = glob.glob(f'{target}/**/*.xlsm', recursive=True)
    xlsx = glob.glob(f'{target}/**/*.xlsx', recursive=True)
    docx = glob.glob(f'{target}/**/*.docx', recursive=True)
    return sorted(xlsm + xlsx + docx)

def ExtractImages(base, path, outdir):
    Log.Info(path)
    zip = zipfile.ZipFile(path)
    file_list = zip.namelist()
    for file in file_list:
        if '/media/' in file:
            img = zip.open(file)
            bin = img.read()
            img_dir = f'{outdir}/{pathlib.Path(path).relative_to(base)}'
            os.makedirs(img_dir, exist_ok=True)
            img_path = f'{img_dir}/{pathlib.Path(file).name}'
            with open(img_path, mode='wb') as f:
                f.write(bin)
            img.close()

def Extract(target, outdir):
    files = EnumOfficeFiles(target)
    for f in files:
        ExtractImages(target, f, outdir)

def Main():
    args = InitArgParser().parse_args()

    if not os.path.exists(args.target):
        Log.Error(f"ディレクトリが見つかりません（{args.target}）")
        return

    Extract(args.target, args.outdir)
    

if __name__ == '__main__':
    Main()