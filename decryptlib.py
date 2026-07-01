"""
Office Open XML Agile Encryption Decryptor (ECMA-376 spec)
Pure Python: pycryptodome + stdlib only, KHÔNG cần msoffcrypto-tool
"""
import io, struct, hashlib, base64, xml.etree.ElementTree as ET
import olefile

try:
    from Crypto.Cipher import AES
except ImportError:
    AES = None

# Standard block keys
BLK_VI = bytes([0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79])
BLK_VH = bytes([0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E])
BLK_EK = bytes([0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6])

NS = 'http://schemas.microsoft.com/office/2006/encryption'
PNS = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password'


def decrypt_agile(file_bytes, password):
    if AES is None:
        raise ImportError("pycryptodome required: pip install pycryptodome")

    ole = olefile.OleFileIO(io.BytesIO(file_bytes))
    enc_info = ole.openstream('EncryptionInfo').read()
    root = ET.fromstring(enc_info[8:])

    kd = root.find(f'./{{{NS}}}keyData')
    data_salt = base64.b64decode(kd.get('saltValue'))

    ek = root.find(f'.//{{{NS}}}keyEncryptors/{{{NS}}}keyEncryptor[@uri="{PNS}"]/{{{PNS}}}encryptedKey')
    spin = int(ek.get('spinCount'))
    ek_salt = base64.b64decode(ek.get('saltValue'))
    verifier_input  = base64.b64decode(ek.get('encryptedVerifierHashInput'))
    verifier_hash   = base64.b64decode(ek.get('encryptedVerifierHashValue'))
    encrypted_key   = base64.b64decode(ek.get('encryptedKeyValue'))

    # --- Step 1: Iterated hash từ password ---
    h = hashlib.sha512(ek_salt + password.encode('utf-16-le'))
    for i in range(0, spin):
        h = hashlib.sha512(struct.pack('<I', i) + h.digest())
    h_digest = h.digest()

    # --- Step 2: Verify password (decrypt cả verifier input + hash value) ---
    k1 = hashlib.sha512(h_digest + BLK_VI).digest()[:32]
    verifier = AES.new(k1, AES.MODE_CBC, iv=ek_salt).decrypt(verifier_input)[:16]

    k2 = hashlib.sha512(h_digest + BLK_VH).digest()[:32]
    decrypted_hash = AES.new(k2, AES.MODE_CBC, iv=ek_salt).decrypt(verifier_hash)[:64]

    if hashlib.sha512(verifier).digest() != decrypted_hash:
        raise ValueError("Wrong password!")

    # --- Step 3: Decrypt secret key ---
    k3 = hashlib.sha512(h_digest + BLK_EK).digest()[:32]
    secret_key = AES.new(k3, AES.MODE_CBC, iv=ek_salt).decrypt(encrypted_key)[:32]

    # --- Step 4: Decrypt EncryptedPackage (4096-byte segments) ---
    enc_pkg = ole.openstream('EncryptedPackage').read()
    total_size = struct.unpack_from('<Q', enc_pkg, 0)[0]

    result = io.BytesIO()
    remaining = total_size
    offset = 8

    for seg_idx in range(0, (total_size + 4095) // 4096):
        seg_size = min(4096, remaining)
        # Round up to 16-byte boundary for AES-CBC
        block_size = 16
        padded_size = ((seg_size + block_size - 1) // block_size) * block_size
        seg_data = enc_pkg[offset:offset + padded_size]

        seg_iv = hashlib.sha512(data_salt + struct.pack('<I', seg_idx)).digest()[:16]
        decrypted = AES.new(secret_key, AES.MODE_CBC, iv=seg_iv).decrypt(seg_data)
        result.write(decrypted[:seg_size])

        remaining -= seg_size
        offset += seg_size
        if remaining <= 0:
            break

    ole.close()
    return result.getvalue()


def decrypt_excel(file_bytes, password='sp'):
    import openpyxl
    try:
        openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True).close()
        return file_bytes
    except:
        pass
    return decrypt_agile(file_bytes, password)


if __name__ == '__main__':
    import sys, openpyxl
    src = sys.argv[1] if len(sys.argv) > 1 else 'test.xlsx'
    dst = sys.argv[2] if len(sys.argv) > 2 else src.replace('.xlsx', '_DECRYPTED.xlsx')
    pw = sys.argv[3] if len(sys.argv) > 3 else 'sp'

    with open(src, 'rb') as f:
        data = f.read()
    decrypted = decrypt_excel(data, pw)
    with open(dst, 'wb') as f:
        f.write(decrypted)

    wb = openpyxl.load_workbook(io.BytesIO(decrypted), data_only=True)
    print(f"✅ Decrypted OK! {len(wb.sheetnames)} sheets → {dst}")
    wb.close()
