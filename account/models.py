from django.db import models
from django.contrib.auth.models import User

class Title(models.Model):
    title = models.CharField(max_length=25,null = False)

    def __str__(self):
        return self.title
    
class Country(models.Model):
    country = models.CharField(max_length=100,null = False)
    code = models.CharField(max_length=3,null = False)
    ph_code = models.CharField(max_length=3,null = False)

    def __str__(self):
        return self.country

class Author(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)  
    title = models.ForeignKey(Title, on_delete=models.CASCADE)
    institution = models.CharField(max_length=255)
    address = models.TextField()
    city = models.CharField(max_length=100)
    state = models.CharField(max_length=100)
    country = models.ForeignKey(Country, on_delete=models.CASCADE)
    latitude = models.FloatField(null=True, blank=True)
    longitude = models.FloatField(null=True, blank=True)
    mobile_no = models.CharField(max_length=25)
    zipcode = models.CharField(max_length=10)
    orcid_id = models.CharField(max_length=19)
    scopus_id = models.CharField(max_length=15)
    is_active = models.BooleanField(default=True)
    is_reviewer = models.BooleanField(default=False)

    def __str__(self):
        return self.user.username

class Editor(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    affliation = models.CharField(max_length=255) 
    mobile_number = models.CharField(max_length=25)
    is_active = models.BooleanField(default=False) 
    
    def __str__(self):
        return self.user.email

class Modes(models.Model):
    name = models.CharField(max_length=100, unique=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name
    
class Question(models.Model):
    question = models.TextField()

    def __str__(self):
        return self.question[:50]  # Return first 50 characters for brevity
    
class FeedbackOptions(models.Model):
    options = models.CharField(max_length=225)
    value = models.IntegerField()
    
    def __str__(self):
        return f"{self.options} ({self.value})"
    
class FeedbackType(models.Model):
    type = models.CharField(max_length=100)

    def __str__(self):
        return self.type

class FeedbackQuestion(models.Model):
    feedback_type = models.ForeignKey(FeedbackType, on_delete=models.CASCADE)
    question = models.ForeignKey(Question, on_delete=models.CASCADE)

    def __str__(self):
        return f"{self.feedback_type} - {self.question.question[:50]}"

class Feedback(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    feedback_type = models.ForeignKey(FeedbackType, on_delete=models.CASCADE)
    assigned_on = models.DateTimeField(auto_now_add=True)
    created_by = models.ForeignKey(User, related_name='created_feedbacks', on_delete=models.CASCADE)
    is_active = models.BooleanField(default=False)
    
    def __str__(self):
        return f"Feedback by {self.created_by.username} to {self.user.username}"

class FeedbackResponse(models.Model):
    feedback = models.ForeignKey(Feedback, on_delete=models.CASCADE)
    question = models.ForeignKey(Question, on_delete=models.CASCADE)
    options = models.ForeignKey(FeedbackOptions, on_delete=models.CASCADE)
    submitted_on = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Response for Ques: {self.question.question[:50]} by User: {self.feedback.user.username}"
# models.py
class LatexDocument(models.Model):
    title = models.CharField(max_length=100)
    tex_file = models.FileField(upload_to='latex_docs/')
    is_active = models.BooleanField(default=False)

    def __str__(self):
        return self.title
