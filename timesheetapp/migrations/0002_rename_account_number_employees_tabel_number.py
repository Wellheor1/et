# Generated by Django 4.0.3 on 2022-04-27 03:11

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('timesheetapp', '0001_initial'),
    ]

    operations = [
        migrations.RenameField(
            model_name='employees',
            old_name='account_number',
            new_name='tabel_number',
        ),
    ]