from django.conf.urls import url

from . import views

app_name='app1'
urlpatterns = [
    url('', views.sender, name = 'sender'),
    ]