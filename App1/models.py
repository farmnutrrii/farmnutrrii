from django.db import models

# Create your models here.
class ClosingStock(models.Model):
    BIRD_TYPES = [
        ('BROILER', 'Broiler'),
        ('CC', 'CC'),
        ('ORIGINAL', 'Original'),
        ('QUAIL', 'Quail'),
    ]
    
    date = models.DateField()
    day = models.CharField(max_length=50)
    bird_type = models.CharField(max_length=10, choices=BIRD_TYPES)
    no_of_birds = models.PositiveIntegerField()
    no_of_kgs = models.DecimalField(max_digits=10, decimal_places=4)
    mortality = models.PositiveIntegerField()

    def __str__(self):
        return f"{self.date} - {self.bird_type}"
    
class DailySheet(models.Model):
    date = models.DateField()
    day = models.CharField(max_length=50)
    bird_type = models.CharField(max_length=50)
    number_of_birds_stock = models.PositiveIntegerField()
    number_of_birds_purchase = models.DecimalField(max_digits=10, decimal_places=4)
    total_birds = models.PositiveIntegerField()
    total_stock_weight = models.DecimalField(max_digits=10, decimal_places=4)
    total_purchase_weight = models.DecimalField(max_digits=10, decimal_places=4)
    total_weight = models.DecimalField(max_digits=10, decimal_places=4)

    def __str__(self):
        return f"{self.bird_type} - {self.id}"

class DailySales(models.Model):
    date = models.DateField()
    day = models.CharField(max_length=50)
    bird_type = models.CharField(max_length=50)
    live_weight = models.DecimalField(max_digits=10, decimal_places=4)
    curry_weight = models.DecimalField(max_digits=10, decimal_places=4)
    day_rate = models.PositiveIntegerField()
    total_sales_amount = models.DecimalField(max_digits=10, decimal_places=4)
    expense = models.DecimalField(max_digits=10, decimal_places=4)
    balance_cash = models.DecimalField(max_digits=10, decimal_places=4)
    gpay = models.DecimalField(max_digits=10, decimal_places=4)

    def __str__(self):
        return f"{self.bird_type} - {self.id}"
    
class WeeklyReport(models.Model):
    date = models.DateField()
    day = models.CharField(max_length=10)
    bird_type = models.CharField(max_length=50)
    number_of_birds = models.IntegerField()
    total_kilograms = models.DecimalField(max_digits=10, decimal_places=4)
    average_weight = models.DecimalField(max_digits=10, decimal_places=4)
    rate = models.DecimalField(max_digits=10, decimal_places=4)
    total_amount = models.DecimalField(max_digits=10, decimal_places=4)
    remarks = models.TextField()

    def __str__(self):
        return f'{self.bird_type} on {self.date}'
