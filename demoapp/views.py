from wsgiref.simple_server import demo_app
from django.shortcuts import render

# Create your views here.
def index(request):
    upload_range = range(1, 13)
    return render(request, 'demoapp/upload_form.html', {'upload_range': upload_range})
