from django.contrib import admin
from .models import ClosingStock, DailySheet, DailySales, WeeklyReport

admin.site.register(ClosingStock)
admin.site.register(DailySheet)
admin.site.register(DailySales)
admin.site.register(WeeklyReport)
