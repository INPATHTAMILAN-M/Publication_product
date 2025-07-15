from django.contrib import admin
from .models import *
# Register your models here.
admin.site.register(Author)
admin.site.register(Title)
admin.site.register(Country)
admin.site.register(Editor)
admin.site.register(Modes)
admin.site.register(Question)
admin.site.register(FeedbackType)
admin.site.register(FeedbackQuestion)
admin.site.register(Feedback)
admin.site.register(FeedbackResponse)
admin.site.register(FeedbackOptions)
admin.site.register(LatexDocument)