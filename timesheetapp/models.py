from django.db import models


class Departments(models.Model):
    title = models.CharField(max_length=255)


class Persons(models.Model):
    last_name = models.CharField(max_length=255)
    first_name = models.CharField(max_length=255)
    patronymic = models.CharField(max_length=255)
    snils = models.CharField(max_length=255)


class TypesPosts(models.Model):
    title = models.CharField(max_length=255)


class Posts(models.Model):
    title = models.CharField(max_length=255)


class Employees(models.Model):
    person = models.ForeignKey(Persons, on_delete=models.CASCADE)
    type_post = models.ForeignKey(TypesPosts, on_delete=models.CASCADE)
    post = models.ForeignKey(Posts, on_delete=models.CASCADE)
    department = models.ForeignKey(Departments, on_delete=models.CASCADE)
    tabel_number = models.DecimalField(max_digits=10, decimal_places=2)


class TimeWork(models.Model):
    employee = models.ForeignKey(Employees, on_delete=models.CASCADE)
    date = models.DateField(auto_now_add=True)
    night_hours = models.DecimalField(max_digits=10, decimal_places=2)
    all_hours = models.DecimalField(max_digits=10, decimal_places=2)

