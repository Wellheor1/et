from django.contrib import admin
from .models import Departments, Persons, TypesPosts, Posts, Employees, TimeWork
# Register your models here.
admin.site.register(Departments)
admin.site.register(Persons)
admin.site.register(TypesPosts)
admin.site.register(Posts)
admin.site.register(Employees)
admin.site.register(TimeWork)
