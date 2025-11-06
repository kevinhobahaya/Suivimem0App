from django.contrib import admin
from .models import EtatSortie

@admin.register(EtatSortie)
class EtatSortieAdmin(admin.ModelAdmin):
    list_display = ('reference','importateur','date','created_at')
    search_fields = ('reference','importateur','container','num_bl')
