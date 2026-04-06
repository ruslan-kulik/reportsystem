from django.db import models


class Category(models.Model):
    name = models.CharField("Название категории", max_length=100, unique=True)

    def __str__(self):
        return self.name

    class Meta:
        verbose_name = "Категория"
        verbose_name_plural = "Категории"


class Product(models.Model):
    name = models.CharField("Название товара", max_length=150, unique=True)
    category = models.ForeignKey(Category, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Категория")
    base_price = models.DecimalField("Базовая стоимость", max_digits=10, decimal_places=2, default=0.00)

    def __str__(self):
        return self.name

    class Meta:
        verbose_name = "Товар"
        verbose_name_plural = "Товары"


class Manager(models.Model):
    full_name = models.CharField("ФИО менеджера", max_length=150, unique=True)

    def __str__(self):
        return self.full_name

    class Meta:
        verbose_name = "Менеджер"
        verbose_name_plural = "Менеджеры"


class Report(models.Model):
    manager = models.ForeignKey(Manager, on_delete=models.CASCADE, verbose_name="Менеджер")
    report_date = models.DateField("Дата отчёта")
    comments = models.TextField("Комментарии", blank=True, null=True)
    created_at = models.DateTimeField("Создан", auto_now_add=True)

    def __str__(self):
        return f"Отчёт {self.manager} от {self.report_date}"

    class Meta:
        verbose_name = "Отчёт"
        verbose_name_plural = "Отчёты"
        ordering = ['-created_at']


class ReportItem(models.Model):
    report = models.ForeignKey(Report, on_delete=models.CASCADE, related_name="items", verbose_name="Отчёт")

    product = models.ForeignKey(Product, on_delete=models.SET_NULL, null=True, blank=True,
                                verbose_name="Товар из списка")
    custom_product_name = models.CharField("Название товара вручную", max_length=150, blank=True, null=True)

    category = models.ForeignKey(Category, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Категория")
    quantity = models.PositiveIntegerField("Количество")
    price_used = models.DecimalField("Цена в отчёте", max_digits=10, decimal_places=2, null=True, blank=True)

    def __str__(self):
        name = self.custom_product_name or (self.product.name if self.product else "Не указан")
        return f"{name} x{self.quantity}"

    class Meta:
        verbose_name = "Позиция отчёта"
        verbose_name_plural = "Позиции отчётов"