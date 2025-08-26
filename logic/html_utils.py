def format_html_table(headers, rows):
    """Форматирует данные в HTML таблицу"""
    html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%; margin: 10px 0;">'
    
    # Заголовки
    html += '<tr style="background-color: #f2f2f2; font-weight: bold;">'
    for header in headers:
        html += f'<th style="border: 1px solid #ddd; padding: 8px; text-align: center;">{header}</th>'
    html += '</tr>'
    
    # Данные
    for row in rows:
        html += '<tr>'
        for cell in row:
            html += f'<td style="border: 1px solid #ddd; padding: 8px;">{cell}</td>'
        html += '</tr>'
    
    html += '</table>'
    return html

