import docx
import json
from tkinter import Tk, filedialog, messagebox


def extract_publications(doc_path):
    doc = docx.Document(doc_path)
    publications = []

    current_publication = None
    body = []

    def parse_meta_data(line):
        """Парсинг мета-даних із тексту."""
        meta = {}
        parts = line.split("\n")
        for part in parts:
            if part.startswith("Час публікації:"):
                meta["Час публікації"] = part.replace("Час публікації:", "").strip()
            elif part.startswith("Джерело:"):
                meta["Джерело"] = part.replace("Джерело:", "").strip()
            elif part.startswith("Тип джерела:"):
                meta["Тип джерела"] = part.replace("Тип джерела:", "").strip()
            elif part.startswith("Адреса оригіналу:"):
                meta["Адреса оригіналу"] = part.replace("Адреса оригіналу:", "").strip()
            elif part.startswith("Охоплення публікації:"):
                meta["Охоплення публікації"] = part.replace("Охоплення публікації:", "").strip()
            elif part.startswith("Адреса в Semantrum:"):
                meta["Адреса в Semantrum"] = part.replace("Адреса в Semantrum:", "").strip()
        return meta

    # Перебір усіх параграфів у документі
    for para in doc.paragraphs:
        text = para.text.strip()

        # Перевірка на початок нової публікації
        if para.style and para.style.name == 'Heading 2':
            # Зберігаємо попередню публікацію, якщо є
            if current_publication:
                current_publication["body"] = "\n".join(body).strip()
                publications.append(current_publication)
            # Початок нової публікації
            current_publication = {
                "title": text,
                "meta_data": {},
                "body": ""
            }
            body = []

        # Якщо текст — мета-дані
        elif current_publication and any(
            key in text for key in ["Час публікації:", "Джерело:", "Тип джерела:", "Адреса оригіналу:", "Охоплення публікації:", "Адреса в Semantrum:"]
        ):
            current_publication["meta_data"].update(parse_meta_data(text))

        # Якщо це текст публікації
        elif current_publication:
            body.append(text)

    # Додаємо останню публікацію
    if current_publication:
        current_publication["body"] = "\n".join(body).strip()
        publications.append(current_publication)

    # Створення JSON
    json_data = {"publications": publications}
    return json.dumps(json_data, ensure_ascii=False, indent=4)


def main():
    # Створюємо вікно для вибору файлу
    root = Tk()
    root.withdraw()  # Ховаємо головне вікно
    root.title("DOCX to JSON Converter")

    # Вибір DOCX файлу
    doc_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if not doc_path:
        messagebox.showinfo("Відміна", "Файл не вибрано")
        return

    try:
        # Виконуємо конвертацію
        json_output = extract_publications(doc_path)
        output_file = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
            initialfile="output1.json"
        )
        if not output_file:
            messagebox.showinfo("Відміна", "Файл для збереження не вибрано")
            return

        # Записуємо результат
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(json_output)

        messagebox.showinfo("Успіх", f"JSON успішно збережено у файлі:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Помилка", f"Сталася помилка:\n{e}")

if __name__ == "__main__":
    main()
