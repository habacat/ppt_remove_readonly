import os
import shutil
import zipfile
import tempfile
import argparse
from lxml import etree
from concurrent.futures import ProcessPoolExecutor, as_completed
import multiprocessing

def remove_modify_verifier(input_pptx, output_pptx):
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(input_pptx, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)

            presentation_xml_path = os.path.join(tmpdir, 'ppt', 'presentation.xml')

            if not os.path.exists(presentation_xml_path):
                return False, "找不到 ppt/presentation.xml 文件。"

            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(presentation_xml_path, parser)
            root = tree.getroot()

            namespaces = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}

            modify_verifiers = root.findall('.//p:modifyVerifier', namespaces)

            if not modify_verifiers:
                pass
            else:
                for verifier in modify_verifiers:
                    parent = verifier.getparent()
                    parent.remove(verifier)

                tree.write(presentation_xml_path, xml_declaration=True, encoding='UTF-8', pretty_print=True)

            with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for foldername, subfolders, filenames in os.walk(tmpdir):
                    for filename in filenames:
                        file_path = os.path.join(foldername, filename)
                        rel_path = os.path.relpath(file_path, tmpdir)
                        zip_out.write(file_path, rel_path)

        return True, None
    except Exception as e:
        return False, str(e)

def process_file(file_path, output_dir):
    try:
        filename = os.path.basename(file_path)
        output_path = os.path.join(output_dir, filename)
        success, error = remove_modify_verifier(file_path, output_path)
        if success:
            return (file_path, True, None)
        else:
            return (file_path, False, error)
    except Exception as e:
        return (file_path, False, str(e))

def main():
    parser = argparse.ArgumentParser(description='将只读 PPTX/PPT 文件转换为可编辑的 PPTX 文件。')
    parser.add_argument('path', nargs='?', help='输入的PPTX/PPT文件或文件夹路径')
    args = parser.parse_args()

    input_folder = 'input'
    output_folder = 'output'

    if os.path.isdir(input_folder):
        input_paths = []
        for root_dir, _, files in os.walk(input_folder):
            for file in files:
                if file.lower().endswith(('.pptx', '.ppt')):
                    input_paths.append(os.path.join(root_dir, file))
    elif args.path:
        if os.path.isdir(args.path):
            input_paths = []
            for root_dir, _, files in os.walk(args.path):
                for file in files:
                    if file.lower().endswith(('.pptx', '.ppt')):
                        input_paths.append(os.path.join(root_dir, file))
        elif os.path.isfile(args.path) and args.path.lower().endswith(('.pptx', '.ppt')):
            input_paths = [args.path]
        else:
            print("错误：指定的路径不是有效的PPTX/PPT文件或文件夹。")
            return
    else:
        print("错误：当前目录下不存在 'input' 文件夹，且未提供输入路径。请指定输入文件或文件夹路径。")
        return

    if not input_paths:
        print("没有找到需要处理的PPTX/PPT文件。")
        return

    os.makedirs(output_folder, exist_ok=True)

    total = len(input_paths)
    success_count = 0
    fail_count = 0
    fail_details = []

    cpu_count = multiprocessing.cpu_count()
    with ProcessPoolExecutor(max_workers=cpu_count) as executor:
        future_to_file = {executor.submit(process_file, file_path, output_folder): file_path for file_path in input_paths}
        for future in as_completed(future_to_file):
            file_path = future_to_file[future]
            try:
                _, success, error = future.result()
                if success:
                    success_count += 1
                else:
                    fail_count += 1
                    fail_details.append((file_path, error))
            except Exception as e:
                fail_count += 1
                fail_details.append((file_path, str(e)))

    print(f"遍历 {total} 个文件，已成功 {success_count} 个文件，失败 {fail_count} 个文件。")
    if fail_count > 0:
        print("失败的文件及原因：")
        for file, reason in fail_details:
            print(f"{file}: {reason}")

if __name__ == "__main__":
    main()
