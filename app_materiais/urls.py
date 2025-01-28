from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_files, name='upload_files'),
    path('download_wb/', views.download_wb, name='download_wb'),
]
