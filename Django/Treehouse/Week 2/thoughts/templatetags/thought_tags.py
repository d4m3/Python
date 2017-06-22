import datetime
from django import template
from django.utils import timezone
import json


# thoughts.forms
from ..forms import ThoughtForm


register = template.Library()

# https://docs.djangoproject.com/en/1.11/howto/custom-template-tags/
@register.inclusion_tag('thoughts/_form.html')
def thought_form():
    form = ThoughtForm()
    return {'form': form }

# Pushing data to chart on dashboard
@register.simple_tag(takes_context=True)
def chart_data(context):
    user = context['user']
    ten_days_ago = timezone.now() - datetime.timedelta(days=10)
    # Change the way the graph is ordered
    thoughts = user.thoughts.filter(recorded_at__gte=ten_days_ago).order_by('recorded_at')
    return json.dumps({
        'labels' : [thought.recorded_at.strftime('%Y-%m-%d') for
                    thought in thoughts],
        # Invert chart with *-1
        'series' : [[thought.condition*-1 for thought in thoughts]]
    })


