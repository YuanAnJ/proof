from django.urls import path

from . import views

app_name = 'forms'
urlpatterns = [
    # Template views
    path('', views.index, name='index'),
    path('form/', views.add_form_template, name='form'),
    path('query/', views.query_form_template, name='query'),
    path('edit/<str:number>', views.edit_form_template, name='edit'),
    path('export/', views.export_form_template, name='export'),
    path('statistic/',views.statistics_form_template,name='statistic'),

    # API endpoints
    path('api/add', views.add_form_api, name='add'),
    path('api/update/<str:number>', views.update_form_api, name='update'),
    path('api/delete/<str:number>', views.delete_form_api, name='delete'),
    path('api/generate/<str:number>', views.generate_doc_api, name='generate'),
    path('api/export-excel/', views.export_excel_api, name='export_excel'),
]

