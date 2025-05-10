from http.server import BaseHTTPRequestHandler
import json
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

class ASSConverter:
    @staticmethod
    def convert(ass_content):
        """Конвертирует ASS-контент в Excel файл"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Субтитры"

        # Заголовки
        headers = ["Время начала", "Имя актера", "Текст"]
        ws.append(headers)

        # Парсинг ASS
        events = []
        in_events = False
        
        for line in ass_content.splitlines():
            line = line.strip()
            if line == "[Events]":
                in_events = True
                continue
            if in_events and line.startswith("Dialogue:"):
                parts = line.split(",", 9)
                start = parts[1].strip()
                actor = parts[4].strip() if len(parts) > 4 else ""
                text = re.sub(r'\{.*?\}', '', parts[9]).replace("\\N", " ")
                ws.append([start, actor, text])

        # Форматирование
        font = Font(size=16)
        alignment = Alignment(wrap_text=True)
        for row in ws.iter_rows():
            for cell in row:
                cell.font = font
                cell.alignment = alignment

        # Автоширина колонок
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 25 
        ws.column_dimensions['C'].width = 50

        # Сохраняем в буфер
        excel_buffer = BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        return excel_buffer

class RequestHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            
            # Получаем файл из multipart/form-data
            boundary = self.headers['Content-Type'].split('=')[1].encode()
            parts = post_data.split(boundary)
            
            file_content = None
            for part in parts:
                if b'filename="' in part:
                    file_content = part.split(b'\r\n\r\n')[1].rstrip(b'\r\n--')
                    break
            
            if not file_content:
                self.send_error(400, "No file uploaded")
                return
            
            # Конвертация
            excel_file = ASSConverter.convert(file_content.decode('utf-8'))
            
            # Отправка Excel-файла
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="subtitles.xlsx"')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(excel_file.getvalue())
            
        except Exception as e:
            self.send_error(500, f"Error: {str(e)}")

def run():
    server = ('', 8000)
    httpd = HTTPServer(server, RequestHandler)
    print("Server running on port 8000...")
    httpd.serve_forever()

if __name__ == '__main__':
    run()
