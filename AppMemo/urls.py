from django.urls import path
from .views import *
from . import views
# AppMemo/urls.py
from SuiviMemo import urls  # ❌ ceci crée une boucle circulaire

app_name='AppMemo'

urlpatterns = [
    path('home',dashboard_concerne, name= 'home'),
    path('usercompte',createUser, name= 'usercompte'),
    path('',loginPage, name= 'login'),
    path('compte',comptePage,name='compte'),
    path('rapport/<uuid:uuid>/', views.memo_interne_view, name='memo_interne'),
    path('concerne',concernePage, name='concerne'),
    path('detail',pageDetail, name='detail'),
    path('loglogin',loginUser,name='loglogin'),
    path('createpageconcerne', CreateConcerne, name='createpageconcerne'),
    path('etatsortie/', views.etat_sortie_list, name='etatsortie'),
    path('etat/<int:pk>/', etat_sortie_detail, name='etat_detail'),
    path('etat/<int:pk>/pdf/', etat_sortie_pdf, name='etat_pdf'),
    path('modifier/<int:pk>/', UpdateConcerne, name='modifier'),
    path('updateD/<int:pk>/', updateConcerneDetail, name='updateD'),
    path('statistiques/', views.tableau_stat_mensuel, name='statistiques'),
    path('export-statistiques-mensuelles/', views.export_stat_mensuel_excel, name='export_stat_mensuel_excel'),
    path('export-statistiques-mensuelles-pdf/', views.export_stat_mensuel_pdf, name='export_stat_mensuel_pdf'),
]
