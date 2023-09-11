from django.contrib import admin
from django.urls import path
from django.contrib.auth.decorators import login_required
from django.contrib.auth import views as auth_views
from . import views
urlpatterns = [
    path('admin/', admin.site.urls),
    path('', login_required(views.home), name="home-page"),
    path('home/', login_required(views.home), name='home'),
    path('success/', login_required(views.home), name='success_page_name'),
    path('register/', views.register, name='register'),
    path('accounts/login/', views.custom_login, name='login'),
]
