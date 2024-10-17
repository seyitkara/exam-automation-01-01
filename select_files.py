from tkinter import Tk, simpledialog, filedialog

def select_files():
    root = Tk()
    root.withdraw()  # Pencereyi gizle

    # Kullanıcıdan birden fazla dosya seçmesini iste
    file_paths = filedialog.askopenfilenames(title="Seçmek istediğiniz dosyaları belirtin")

    # Seçilen dosya yollarını yazdır (Burayı istediğiniz gibi değiştirebilirsiniz)
    for file in file_paths:
        print(file)

    # Tkinter penceresini kapat
    root.destroy()

# Fonksiyonu diğer dosyada çağır