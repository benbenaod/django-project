from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path("", include("main.urls")),   # 首頁走 main.urls
    path("admin/", admin.site.urls),
]
