import word_writer

if word_writer.com_available():
    print("Word COM: CO SAN - se dung che do Word COM (do trung thuc cao nhat).")
else:
    print(
        "Word COM: KHONG CO - se dung che do python-docx thuan (fallback).\n"
        "Luu y: mot vai cho xoa dong trong file '00. QD Giao de tai' co the "
        "de sot dong trong, script se canh bao khi gap."
    )
