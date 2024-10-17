from tkinter import Tk, simpledialog

def get_output_filename():
    root = Tk()
    root.withdraw()  # Pencereyi gizle

    # Kullanıcıdan çıktı dosya ismi iste
    output_filename = simpledialog.askstring(title="Çıktı Dosya İsmi",
                                           prompt="Çıktı Excel dosyasının adını giriniz (örneğin: exam-results.xlsx):")

    # Girilen ismi döndür
    return output_filename
