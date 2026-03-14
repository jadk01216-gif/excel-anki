from converter import AnkiConverter
import os

def test_conversion():
    excel_path = r"c:\Users\USER\Downloads\english learning\2026 new1  (1).xlsx"
    output_path = r"c:\Users\USER\Downloads\english learning\test_output.apkg"
    
    if not os.path.exists(excel_path):
        print(f"Error: {excel_path} not found.")
        return

    print(f"Starting test conversion for {excel_path}...")
    converter = AnkiConverter(excel_path, output_path, include_tts=True)
    
    # Simple progress callback
    def progress(p):
        print(f"Progress: {p}%")
        
    success = converter.process(progress)
    
    if success and os.path.exists(output_path):
        print(f"Success! Generated {output_path}")
        print(f"File size: {os.path.getsize(output_path)} bytes")
    else:
        print("Conversion failed.")

if __name__ == "__main__":
    test_conversion()
