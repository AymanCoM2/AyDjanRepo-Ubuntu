from django.contrib import admin
from django.urls import path
from django.contrib.auth.decorators import login_required
from django.contrib.auth import views as auth_views
from . import views

# urlpatterns = [
#     path('admin/', admin.site.urls),
#     path('', login_required(views.home), name="home-page"),
#     path('home/', login_required(views.home), name='home'),
#     path('success/', login_required(views.home), name='success_page_name'),
#     path('register/', views.register, name='register'),
#     path('accounts/login/', views.custom_login, name='login'),
# ]

urlpatterns = [
    path('admin/', admin.site.urls),  # ! Ok
    path('', views.main_page, name="home_page"),  # ! Ok
    path('main-page/', views.main_page, name="main_page"),  # ! Ok
    path('card-code/', views.cardcodeController, name='card_code'),  # ! Ok
    path('card-code-b/', views.cardcodeControllerB, name='card_codeB'),  # ! Ok
    path('all-code/', views.allcodeController, name='all_code'),  # ^ NOT Ok
    path('success/', views.main_page, name='success_page_name'),  # ! Ok
    path('register/', views.register, name='register'),  # ! Ok
    path('accounts/login/', views.custom_login, name='login'),  # ! Ok
]
