from django.urls import path
from . import views

urlpatterns = [
    path('', views.report_list, name='report_list'),
    path('create/', views.report_create, name='report_create'),
    path('admin-panel/', views.admin_panel, name='admin_panel'),
    path('search/', views.search_reports, name='search_reports'),
    path('manager/<int:manager_id>/', views.manager_detail, name='manager_detail'),
    path('export/pdf/<int:report_id>/', views.export_pdf, name='export_pdf'),
    path('export/xml/<int:report_id>/', views.export_xml, name='export_xml'),
    path('export/xlsx/<int:report_id>/', views.export_xlsx, name='export_xlsx'),
]