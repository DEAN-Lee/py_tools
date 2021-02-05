from django.urls import path

from . import views

urlpatterns = [
    path(r'users', views.UserViewSet),
    path(r'groups', views.GroupViewSet)
]
