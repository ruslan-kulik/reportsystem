from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse
from django.contrib import messages
from django.db.models import Q
from io import BytesIO
import xml.etree.ElementTree as ET

from io import BytesIO
from django.shortcuts import get_object_or_404, render
from django.http import HttpResponse
from reportlab.lib.fonts import addMapping
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from io import BytesIO
from django.shortcuts import get_object_or_404
from django.http import HttpResponse
import os

from .models import Report, Manager, Category, Product, ReportItem
from .forms import (
    ManagerForm, CategoryForm, ProductForm,
    ReportForm, ReportItemFormSet, SearchForm
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from datetime import datetime

def admin_panel(request):
    if request.method == 'POST':
        if 'manager_submit' in request.POST:
            form = ManagerForm(request.POST)
            if form.is_valid():
                form.save()
                messages.success(request, "Менеджер добавлен")
        elif 'category_submit' in request.POST:
            form = CategoryForm(request.POST)
            if form.is_valid():
                form.save()
                messages.success(request, "Категория добавлена")
        elif 'product_submit' in request.POST:
            form = ProductForm(request.POST)
            if form.is_valid():
                form.save()
                messages.success(request, "Товар добавлен")
        return redirect('admin_panel')

    return render(request, 'reports/admin_panel.html', {
        'managers': Manager.objects.all(),
        'categories': Category.objects.all(),
        'products': Product.objects.all(),
        'manager_form': ManagerForm(),
        'category_form': CategoryForm(),
        'product_form': ProductForm()
    })


def report_list(request):
    reports = Report.objects.select_related('manager').all()
    return render(request, 'reports/report_list.html', {'reports': reports})


def manager_detail(request, manager_id):
    manager = get_object_or_404(Manager, pk=manager_id)
    reports = Report.objects.filter(manager=manager).order_by('-report_date')
    return render(request, 'reports/manager_detail.html', {'manager': manager, 'reports': reports})


def report_create(request):
    if request.method == 'POST':
        report_form = ReportForm(request.POST)
        item_formset = ReportItemFormSet(request.POST)

        if report_form.is_valid() and item_formset.is_valid():
            report = report_form.save()
            items = item_formset.save(commit=False)
            for item in items:
                item.report = report
                if not item.price_used and item.product:
                    item.price_used = item.product.base_price
                if item.custom_product_name and item.product:
                    item.product = None
                item.save()
            messages.success(request, "Отчёт успешно создан!")
            return redirect('report_list')
    else:
        report_form = ReportForm()
        item_formset = ReportItemFormSet()

    return render(request, 'reports/report_create.html', {
        'report_form': report_form,
        'item_formset': item_formset
    })


def search_reports(request):
    form = SearchForm(request.GET or None)
    reports = Report.objects.select_related('manager').prefetch_related('items').all()

    if form.is_valid():
        cd = form.cleaned_data
        if cd['date_from']:
            reports = reports.filter(report_date__gte=cd['date_from'])
        if cd['date_to']:
            reports = reports.filter(report_date__lte=cd['date_to'])
        if cd['manager']:
            reports = reports.filter(manager=cd['manager'])
        if cd['keyword']:
            kw = cd['keyword']
            q_objects = Q(comments__icontains=kw) | Q(manager__full_name__icontains=kw)
            q_objects |= Q(items__product__name__icontains=kw) | Q(items__custom_product_name__icontains=kw)
            reports = reports.filter(q_objects).distinct()

    return render(request, 'reports/search.html', {'form': form, 'reports': reports})


# ================= ЭКСПОРТ PDF (с регистрацией шрифта Arial) =================
def export_pdf(request, report_id):
    report = get_object_or_404(Report, pk=report_id)

    # 1. Регистрация шрифта Arial
    arial_path = "C:/Windows/Fonts/arial.ttf"
    if not os.path.exists(arial_path):
        return HttpResponse("Ошибка: Шрифт Arial не найден в системе.", status=500)

    try:
        pdfmetrics.registerFont(TTFont('Arial', arial_path))
    except:
        pass  # Если уже зарегистрирован

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=2 * cm, leftMargin=2 * cm, topMargin=2 * cm,
                            bottomMargin=2 * cm)
    elements = []

    # Стили
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(name='Title', fontName='Arial', fontSize=16, alignment=1, spaceAfter=30)
    style_normal = ParagraphStyle(name='Normal', fontName='Arial', fontSize=12, leading=14)
    style_bold = ParagraphStyle(name='Bold', fontName='Arial', fontSize=12, fontWeight='bold')

    # Заголовок
    elements.append(Paragraph(f"Отчёт о продажах №{report.id}", style_title))

    # Инфо
    elements.append(Paragraph(f"<b>Менеджер:</b> {report.manager.full_name}", style_normal))
    elements.append(Paragraph(f"<b>Дата:</b> {report.report_date}", style_normal))
    if report.comments:
        elements.append(Paragraph(f"<b>Комментарий:</b> {report.comments}", style_normal))

    elements.append(Spacer(1, 0.5 * cm))

    # Таблица данных
    data = [['Товар', 'Категория', 'Кол-во', 'Цена', 'Сумма']]

    for item in report.items.all():
        name = item.custom_product_name or (item.product.name if item.product else "-")
        cat = item.category.name if item.category else "-"
        qty = str(item.quantity)
        price = f"{float(item.price_used or 0):.2f}"
        total = f"{(float(item.price_used or 0) * item.quantity):.2f}"
        data.append([name, cat, qty, price, total])

    table = Table(data, colWidths=[8 * cm, 3 * cm, 2 * cm, 2 * cm, 2 * cm])
    table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BACKGROUND', (0, 0), (-1, 0), '#eeeeee'),
        ('GRID', (0, 0), (-1, -1), 1, '#000000'),
        ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
    ]))

    elements.append(table)

    doc.build(elements)

    response = HttpResponse(buffer.getvalue(), content_type='application/pdf')
    response['Content-Disposition'] = f'inline; filename="report_{report.id}.pdf"'
    return response
def export_xml(request, report_id):
    report = get_object_or_404(Report, pk=report_id)
    root = ET.Element("Report", id=str(report.id), manager=report.manager.full_name, date=str(report.report_date))

    for item in report.items.all():
        item_el = ET.SubElement(root, "Item")
        name = item.custom_product_name or (item.product.name if item.product else "Неизвестно")
        ET.SubElement(item_el, "Name").text = name
        ET.SubElement(item_el, "Category").text = item.category.name if item.category else "Не указана"
        ET.SubElement(item_el, "Quantity").text = str(item.quantity)
        ET.SubElement(item_el, "Price").text = str(item.price_used or 0)
        total = (item.price_used or 0) * item.quantity
        ET.SubElement(item_el, "Total").text = str(total)

    xml_str = ET.tostring(root, encoding='unicode', method='xml')
    xml_response = '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_str

    response = HttpResponse(xml_response, content_type='text/xml; charset=utf-8')
    response['Content-Disposition'] = f'attachment; filename="report_{report.id}.xml"'
    return response





def export_xlsx(request, report_id):
    report = get_object_or_404(Report, pk=report_id)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Отчет {report.id}"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    ws.merge_cells('A1:E1')
    cell_title = ws['A1']
    cell_title.value = f"ОТЧЁТ О ПРОДАЖАХ №{report.id}"
    cell_title.font = Font(bold=True, size=14)
    cell_title.alignment = Alignment(horizontal="center")

    ws.append([])

    ws.append(["Менеджер:", report.manager.full_name])
    ws.append(["Дата отчёта:", report.report_date])
    if report.comments:
        ws.append(["Комментарий:", report.comments])

    ws.append([])

    headers = ["№", "Наименование товара", "Категория", "Количество", "Цена (BYN)", "Сумма (BYN)"]
    ws.append(headers)

    for cell in ws[ws.max_row]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_align

    row_num = ws.max_row + 1
    total_sum = 0

    for idx, item in enumerate(report.items.all(), 1):
        name = item.custom_product_name or (item.product.name if item.product else "Не указано")
        category = item.category.name if item.category else "-"
        qty = item.quantity
        price = float(item.price_used or 0)
        item_sum = qty * price
        total_sum += item_sum

        ws.append([idx, name, category, qty, price, item_sum])

        for cell in ws[ws.max_row]:
            cell.border = thin_border
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'  # Формат числа с копейками
                cell.alignment = Alignment(horizontal="right")

    ws.append(["", "", "", "ИТОГО:", "", total_sum])
    total_row = ws[ws.max_row]
    total_row[3].font = Font(bold=True)
    total_row[5].font = Font(bold=True)
    total_row[5].number_format = '#,##0.00'

    column_widths = [5, 40, 20, 10, 15, 15]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64 + i)].width = width

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = f'attachment; filename=report_{report.id}.xlsx'
    wb.save(response)
    return response