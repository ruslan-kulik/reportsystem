from django import forms
from .models import Product, Category, Manager, Report, ReportItem



class ManagerForm(forms.ModelForm):
    class Meta:
        model = Manager
        fields = ['full_name']
        widgets = {'full_name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Иванов И.И.'})}

class CategoryForm(forms.ModelForm):
    class Meta:
        model = Category
        fields = ['name']
        widgets = {'name': forms.TextInput(attrs={'class': 'form-control'})}

class ProductForm(forms.ModelForm):
    class Meta:
        model = Product
        fields = ['name', 'category', 'base_price']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'category': forms.Select(attrs={'class': 'form-select'}),
            'base_price': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01'})
        }

class ReportForm(forms.ModelForm):
    class Meta:
        model = Report
        fields = ['manager', 'report_date', 'comments']
        widgets = {
            'manager': forms.Select(attrs={'class': 'form-select'}),
            'report_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'comments': forms.Textarea(attrs={'rows': 3, 'class': 'form-control'})
        }

ReportItemFormSet = forms.inlineformset_factory(
    Report,
    ReportItem,
    fields=['product', 'custom_product_name', 'category', 'quantity', 'price_used'],
    extra=1,
    can_delete=True,
    widgets={
        'product': forms.Select(attrs={'class': 'form-select item-product'}),
        'custom_product_name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Или введите название вручную'}),
        'category': forms.Select(attrs={'class': 'form-select item-category'}),
        'quantity': forms.NumberInput(attrs={'class': 'form-control item-qty', 'min': '1', 'value': '1'}),
        'price_used': forms.NumberInput(attrs={'class': 'form-control item-price', 'step': '0.01', 'placeholder': 'Авто'})
    }
)

class SearchForm(forms.Form):
    date_from = forms.DateField(required=False, widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}))
    date_to = forms.DateField(required=False, widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}))
    manager = forms.ModelChoiceField(queryset=Manager.objects.all(), required=False, empty_label="Все менеджеры", widget=forms.Select(attrs={'class': 'form-select'}))
    keyword = forms.CharField(required=False, max_length=200, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Поиск по тексту (ФИО, товар, комментарий...)'}))