import zipfile, os, sys

f = '/Users/Apple/Downloads/lili_work/work1/procalc5/result.xlsx'
print(f"File size: {os.path.getsize(f)} bytes")

try:
    with zipfile.ZipFile(f) as z:
        print("Files in xlsx:")
        for n in z.namelist():
            info = z.getinfo(n)
            print(f"  {n}: {info.file_size} bytes")
        
        for name in z.namelist():
            if 'sheet' in name.lower() and name.endswith('.xml'):
                print(f"\nReading {name}:")
                data = z.read(name).decode('utf-8', errors='replace')
                print(data[:3000])
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
