import os
try:
    from PIL import Image
    
    def convert_to_ico(input_path, output_path):
        img = Image.open(input_path)
        # Standard Windows ICO sizes including the high-res 256px layer for installer and desk icons
        icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        img.save(output_path, sizes=icon_sizes)
        print(f"Successfully converted {input_path} to {output_path}")

    if __name__ == "__main__":
        src = r"c:\Users\Vinrox\Desktop\project_traceability_system\client\src\assets\installer_icon.png"
        dst = r"c:\Users\Vinrox\Desktop\project_traceability_system\electron\installer_icon.ico"
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        
        convert_to_ico(src, dst)
except ImportError:
    print("Error: PIL (Pillow) is not installed. Please run 'pip install pillow'.")
except Exception as e:
    print(f"An error occurred: {e}")
