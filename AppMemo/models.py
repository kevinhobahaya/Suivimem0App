from django.db import models
from decimal import Decimal
from django.contrib.auth.models import User
import uuid

  
class Concerne(models.Model):  
    importateur = models.CharField(max_length=255)  
    container = models.CharField(max_length=100)  
    marchandise = models.CharField(max_length=255)  
    quantite = models.PositiveIntegerField()  
    poids = models.DecimalField(max_digits=10, decimal_places=2)  
    fob = models.CharField(max_length=100, default=1)  
    cif =  models.CharField(max_length=255,default=1) 
    num_bl = models.CharField(max_length=100)  
    num_fbr_aa = models.CharField(max_length=100)  
    plaque = models.CharField(max_length=50)  
    num_e = models.CharField(max_length=100)  
    origine = models.CharField(max_length=255)  
    provenance = models.CharField(max_length=255)  
    transitaire = models.CharField(max_length=255) 
    num_ref = models.CharField(max_length=100, default=1)  
    mois_ref = models.CharField(max_length=50, default=1)  
    date = models.DateField(auto_now= True) 
    frais_controle = models.DecimalField(max_digits=10, decimal_places=2)
    frais_analyse_labo = models.DecimalField(max_digits=10, decimal_places=2)
    tva = models.DecimalField(max_digits=10, decimal_places=2)
    
    total_brut = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    nb_no = models.CharField(max_length=100)
    bv_num = models.CharField(max_length=100, default=1)
    recu_no = models.CharField(max_length=100)
    liquidation_no = models.CharField(max_length=100)
    travail_laboratoire = models.CharField(max_length=255)
    destination = models.CharField(max_length=255, default=1) 
    num_av = models.CharField(max_length=255, default=1) 

    # 👇 Ajout du champ UUID
    uuid = models.UUIDField(default=uuid.uuid4, editable=False, unique=True, )
   
  
    def __str__(self):  
        return f"{self.importateur} - {self.marchandise}"  

class Profil(models.Model):
    name = models.CharField(max_length=100, unique=True)
    desc = models.ManyToManyField(User, related_name='desc') 

class Role(models.Model):
    name = models.CharField(max_length=100, unique=True)
    users = models.ManyToManyField(User, related_name='roles')
    
   
    def __str__(self):
        return self.name

class EtatSortie(models.Model):
    date = models.DateField(null=True, blank=True)
    reference = models.CharField(max_length=100, blank=True)
    importateur = models.CharField(max_length=200, blank=True)
    container = models.CharField(max_length=100, blank=True)
    marchandises = models.TextField(blank=True)
    quantite = models.CharField(max_length=100, blank=True)
    poids = models.CharField(max_length=100, blank=True)
    fob = models.CharField(max_length=100, blank=True)
    cif = models.CharField(max_length=100, blank=True)
    num_bl = models.CharField(max_length=100, blank=True)
    num_feriad = models.CharField("Num. FERI/AD", max_length=100, blank=True)
    plaque = models.CharField(max_length=100, blank=True)
    num_e = models.CharField("Num. E", max_length=100, blank=True)
    origine = models.CharField(max_length=100, blank=True)
    provenance = models.CharField(max_length=100, blank=True)
    transitaire = models.CharField(max_length=200, blank=True)

    # Frais perçus
    frais_controle = models.CharField(max_length=100, blank=True)
    frais_analyse = models.CharField(max_length=100, blank=True)
    tva = models.CharField(max_length=100, blank=True)
    total_brut = models.CharField(max_length=100, blank=True)
    nd_num = models.CharField("N.D. n°", max_length=100, blank=True)
    recu_num = models.CharField("Reçu n°", max_length=100, blank=True)
    bv_num = models.CharField("B.V n°", max_length=100, blank=True)
    liquidation_num = models.CharField("Liquidation n°", max_length=100, blank=True)
    produit_labo = models.CharField("Produit/Laboratoire", max_length=200, blank=True)

    notes = models.TextField(blank=True)

    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"EtatSortie {self.reference or self.id} - {self.importateur}"
