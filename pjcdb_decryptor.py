import win32com.client.dynamic
import os
import time
import sys

def display_banner():
    """
    显示工具的 ASCII 艺术横幅。
    """
    banner = """
 ____     _  ____ ____  _       ____                             _             
|  _ \   | |/ ___|  _ \| |__   |  _ \  ___  ___ _ __ _   _ _ __ | |_ ___  _ __ 
| |_) |  | | |   | | | | \'_ \  | | | |/ _ \/ __| \'_| | | | \'_ \| __/ _ \| \'__|
|  __/ |_| | |___| |_| | |_) | | |_| |  __/ (__| |  | |_| | |_) | || (_) | |   
|_|   \___/ \____|____/|_.__/  |____/ \___|\___|_|   \__, | .__/ \__\___|_|   
                                                     |___/|_|                  
    """
    print(banner)
    print("\n[+] PJCDb Decryptor - Office Document Encryption Bypass Tool")
    print("[+] Version: 1.0.0")
    print("[+] Author: Manus AI (Original by eninem123)")
    print("[+] Disclaimer: For educational and research purposes only. Use responsibly.\n")

def decrypt_office_documents(target_directory="."):
    """
    扫描指定目录下的 PowerPoint 文件，尝试通过另存为操作绕过加密。
    解密后的文件将保存在 'decrypted_output' 文件夹中。
    
    Args:
        target_directory (str): 待扫描的目录路径。
    """
    print(f"[*] Initializing decryption sequence in: {os.path.abspath(target_directory)}")
    
    output_dir_name = "decrypted_output"
    output_path = os.path.join(target_directory, output_dir_name)

    if not os.path.isdir(output_path):
        print(f"[*] Creating output directory: {output_path}")
        os.mkdir(output_path)
    else:
        print(f"[*] Output directory already exists: {output_path}")

    try:
        # 创建 PowerPoint 应用程序对象
        # WithWindow=0 表示在后台运行，不显示界面
        powerpoint_app = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint_app.Visible = False # 确保应用程序不可见
        print("[+] PowerPoint Application instance created successfully.")
    except Exception as e:
        print(f"[CRITICAL] Failed to create PowerPoint Application instance. Error: {e}")
        print("[CRITICAL] Ensure Microsoft Office is installed and accessible.")
        sys.exit(1)

    processed_count = 0
    for root, _, files in os.walk(target_directory):
        for file_name in files:
            full_file_path = os.path.join(root, file_name)
            
            # 检查文件是否为 PowerPoint 格式
            if file_name.lower().endswith(('.pptx', '.ppt')):
                print(f"\n[*] Detected PowerPoint file: {full_file_path}")
                
                decrypted_file_name = os.path.splitext(file_name)[0] + ".ini" # 临时保存为 .ini
                final_decrypted_path = os.path.join(output_path, os.path.splitext(file_name)[0] + os.path.splitext(file_name)[1])
                temp_ini_path = os.path.join(output_path, decrypted_file_name)

                if os.path.exists(final_decrypted_path):
                    print(f"[SKIP] File already decrypted and exists in output: {final_decrypted_path}")
                    continue
                
                print(f"[*] Attempting to bypass encryption for: {file_name}")
                try:
                    # 尝试打开演示文稿
                    presentation = powerpoint_app.Presentations.Open(full_file_path, WithWindow=0)
                    print(f"[+] Successfully opened presentation: {file_name}")
                    
                    # 另存为 .ini 文件，此步骤通常会绕过加密
                    presentation.SaveAs(temp_ini_path) 
                    print(f"[+] Saved temporary decrypted file as: {temp_ini_path}")
                    
                    presentation.Close()
                    print(f"[+] Closed presentation: {file_name}")
                    
                    # 将 .ini 文件重命名回原始扩展名
                    os.rename(temp_ini_path, final_decrypted_path)
                    print(f"[SUCCESS] Decryption bypass successful. Saved to: {final_decrypted_path}")
                    processed_count += 1
                except Exception as e:
                    print(f"[ERROR] Failed to process {file_name}. Error: {e}")
                finally:
                    # 确保即使出错也能关闭演示文稿
                    if 'presentation' in locals() and presentation.Saved == False:
                        presentation.Close()

    # 退出 PowerPoint 应用程序
    powerpoint_app.Quit()
    print("\n[*] PowerPoint Application instance terminated.")
    print(f"[SUMMARY] Decryption process completed. Total files processed: {processed_count}")
    print(f"[SUMMARY] Decrypted files are located in: {os.path.abspath(output_path)}")

if __name__ == "__main__":
    display_banner()
    decrypt_office_documents()
