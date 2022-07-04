from django.templatetags.static import static
from django.urls import reverse
from jinja2 import Environment
from django.contrib.auth.models import Group


def environment(**options):
    env = Environment(**options)
    env.globals.update({'static': static, 'url': reverse, 'groups': Group.objects.all()})
    return env
