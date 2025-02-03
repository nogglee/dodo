import msoffcrypto
import sys

input_file = sys.argv[1]  # 암호화된 원본 파일
output_file = sys.argv[2]  # 암호 해제 후 저장할 파일
password = "1928103362"  # 엑셀 암호

try:
    with open(input_file, "rb") as f:
        encrypted = msoffcrypto.OfficeFile(f)
        encrypted.load_key(password=password)
        with open(output_file, "wb") as decrypted_file:
            encrypted.decrypt(decrypted_file)

    print("SUCCESS")
except Exception as e:
    print("ERROR:", str(e))