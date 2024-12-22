from django.urls import path
from .import views

urlpatterns = [
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('', views.index, name='index'), 
    path('closing-stock/', views.closing_stock_view, name='closing_stock'),
    path('view-stock/<str:stock_date>/', views.view_stock, name='view_stock'),
    path('download_excel_closingstock/', views.download_excel_closingstock, name='download_excel_closingstock'),
    path('dailysheet/', views.dailysheet, name='dailysheet'),
    path('view-stock-dailysheet/<str:stock_date>/', views.view_stock_dailysheet, name='view_stock_dailysheet'),
    path('download_excel_dailysheet/', views.download_excel_dailysheet, name='download_excel_dailysheet'),
    path('daily-sales/', views.daily_sales, name='daily_sales'),
    path('view-stock-dailysales/<str:stock_date>/', views.view_stock_dailysales, name='view_stock_dailysales'),
    path('download_excel_dailysales/', views.download_excel_dailysales, name='download_excel_dailysales'),
    path('weekly-report/', views.weekly_report, name='weekly_report'),
    path('view-stock-weeklyreport/<str:stock_date>/', views.view_stock_weeklyreport, name='view_stock_weeklyreport'),
    path('download_excel/', views.download_excel, name='download_excel'),
]