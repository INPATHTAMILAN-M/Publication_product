from django.db import models
from account.models import Author,Editor
from django.contrib.auth.models import User
from django.utils.timezone import now

class Journal(models.Model):
    title = models.CharField(max_length=255)
    def __str__(self):
        return self.title

class Category(models.Model):
    catagory =models.CharField(max_length=50)
    
    def __str__(self):
        return self.catagory
    
class Article_Type(models.Model):
    article_type = models.CharField(max_length=50)
    article_description = models.CharField(max_length=255)
    
    def __str__(self):
        return self.article_type

class Article_Status(models.Model):
    article_status = models.CharField(max_length=50)
    description = models.TextField()

    def __str__(self):
        return self.article_status
    
class Specialization(models.Model):
    specialization = models.CharField(max_length=50)

    def __str__(self):
        return self.specialization

class Decision(models.Model):
    decision = models.CharField(max_length=50, null=False)
    description = models.TextField(null=True, blank=True)
    
    def __str__(self):
        return f'Decision: {self.decision}'

class File_Category(models.Model):
    file_category = models.CharField(max_length=50, null=False)

    def __str__(self):
        return self.file_category
    
class Request_Status(models.Model):
    request_status = models.CharField(max_length=50)
    
    def __str__(self):
        return self.request_status
    
class Reviewer_Specialization(models.Model):
    reviewer = models.ForeignKey(User, on_delete=models.CASCADE)
    specialization = models.ForeignKey(Specialization, on_delete=models.CASCADE)
    
    def __str__(self):
        return f'{self.reviewer.user} - {self.specialization.specialization}'
 
class Submission(models.Model):
    journal = models.ForeignKey(Journal, on_delete=models.CASCADE, null=True, blank=True)
    author = models.ForeignKey(Author, on_delete=models.CASCADE, null=True, blank=True)
    title = models.CharField(max_length=255, null=True, blank=True)
    manuscript_id = models.CharField(max_length=255, null=True, blank=True)
    article_type = models.ForeignKey(Article_Type, on_delete=models.CASCADE, null=True, blank=True)
    category = models.ForeignKey(Category, on_delete=models.CASCADE, null=True, blank=True)
    abstract = models.TextField(null=True, blank=True)
    is_funded = models.BooleanField(null=True, blank=True)
    no_of_figures = models.IntegerField(null=True, blank=True)
    no_of_tables = models.IntegerField(null=True, blank=True)
    no_of_words = models.IntegerField(null=True, blank=True)
    is_submitted_already = models.BooleanField(null=True, blank=True)
    acknowledgement_1 = models.BooleanField(default=False)
    acknowledgement_2 = models.BooleanField(default=False)
    acknowledgement_3 = models.BooleanField(default=False)
    conflict_of_interest = models.BooleanField(null=True, blank=True)
    coi_describe = models.TextField(null=True, blank=True)
    is_verified = models.BooleanField(default=False)
    cover_letter = models.TextField( null=True, blank=True)   #upload_to='submissions/',
    parent_submission = models.ForeignKey('self', on_delete=models.CASCADE, null=True, blank=True, related_name='patent_submissions')
    submitted_on = models.DateField(auto_now_add=True, null=True, blank=True)
    decissioned_on = models.DateField(auto_now_add=True,null=True, blank=True)
    decision_by = models.ForeignKey(User, on_delete=models.CASCADE,related_name="user", null=True, blank=True)
    article_status = models.ForeignKey(Article_Status, on_delete=models.CASCADE, null=True, blank=True)
    is_decissioned = models.BooleanField(default=False)
    decision = models.ForeignKey(Decision, on_delete=models.CASCADE, null=True, blank=True)
    final_file = models.FileField(upload_to='submissions/', null=True, blank=True)
    admin_comments = models.TextField(null=True, blank=True)
    payment_date = models.DateField(null=True, blank=True)
    payment_due_date = models.DateField(null=True, blank=True)
    revision_due_date = models.DateField(null=True, blank=True)
    typeset_due_date = models.DateField(null=True, blank=True)
    correction_due_date = models.DateField(null=True, blank=True)
    plag_report = models.FileField(upload_to='submissions/', null=True, blank=True)
    copyright_file = models.FileField(upload_to='submissions/', null=True, blank=True)
    additional_file = models.FileField(upload_to='submissions/', null=True, blank=True)
    specialization = models.ForeignKey(Specialization, on_delete=models.CASCADE, null=True, blank=True)
    def __str__(self):
        return self.title or "No Title"


class Correction_Comments(models.Model):
    submission = models.ForeignKey (Submission, on_delete=models.CASCADE,  related_name='correction_comments')
    additional_file = models.FileField(upload_to='submissions/', null=True, blank=True)
    correction_commments = models.TextField()
    date = models.DateField(auto_now_add=True)
    due_date = models.DateField()

    def __str__(self):
        return self.submission.title

class Submission_Files(models.Model):
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE)
    file_category = models.ForeignKey(File_Category, on_delete=models.CASCADE)
    file = models.FileField(upload_to='submissions/')
    file_size = models.FloatField()

    def __str__(self):
        return f"{self.submission.title} - {self.file.name}"

class Submission_Reviewer(models.Model):
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE)
    reviewer = models.ForeignKey(Author, on_delete=models.CASCADE)
    completion_on = models.DateTimeField(null=True, blank=True)
    accepted_on = models.DateTimeField(auto_now_add=True,null=True, blank=True)
    request_status = models.ForeignKey(Request_Status, on_delete=models.CASCADE)
    REVIEW_RECOMMENDATION_CHOICES = [
        ('A', 'Accept'),
        ('R', 'Reject'),
        ('MIN_R', 'Minor Revision'),
        ('MAJ_R', 'Major Revision'),
    ]
    review_recommendation = models.CharField(max_length=5, choices=REVIEW_RECOMMENDATION_CHOICES)
    review_comments = models.TextField()
    due_date = models.DateField(null=True,blank=True)

    def __str__(self):
        return f"{self.reviewer.user.first_name} - {self.submission.title}"
    
class Reviewer_Invitation(models.Model):
    user= models.ForeignKey(User, on_delete=models.CASCADE,related_name= 'reviewer')
    INVITE_STATUS_CHOICES = [
        ('R', 'Requested'),
        ('A', 'Accepted'),
        ('RJ', 'Reject')
    ]
    invite_status = models.CharField(max_length=10, choices=INVITE_STATUS_CHOICES)
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE)
    date = models.DateField(auto_now_add=True)
    expiring_date = models.DateField()
    invite_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='invite_by', null=True, blank=True)
    
    def __str__(self):
        return f"{self.user.first_name} - {self.submission.title}"

class Communication(models.Model):
    title = models.CharField(max_length=255, null=False)
    sender = models.ForeignKey(User, on_delete=models.CASCADE,related_name="sender")
    receiver = models.ForeignKey(User, on_delete=models.CASCADE,related_name="reciever")
    date = models.DateField(auto_now_add=True)
    time = models.TimeField(auto_now_add=True)
    detail = models.TextField(blank=True)
    file = models.FileField(upload_to='communications/', blank=True, null=True)
    submission = models.ForeignKey(Submission , on_delete=models.CASCADE, null=True)

    def __str__(self):
        return self.title

class AE_Assignment(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE)
    date = models.DateField(auto_now_add=True,null=True, blank=True)
    REVIEW_RECOMMENDATION_CHOICES = [
        ('A', 'Accept'),
        ('R', 'Reject'),
        ('MIN_R', 'Minor Revision'),
        ('MAJ_R', 'Major Revision'),
    ]
    ae_recommendation = models.CharField(max_length=5, choices=REVIEW_RECOMMENDATION_CHOICES)
    ae_comments = models.TextField()

    def __str__(self):
        return f"{self.user} - {self.submission}"
    
class Journal_Editor_Assignment(models.Model):
    journal = models.ForeignKey(Journal, on_delete=models.CASCADE)
    editor = models.ForeignKey(Editor, on_delete=models.CASCADE)
    assigned_date = models.DateField(auto_now_add=True)

    def __str__(self):
        return f"{self.journal} - {self.editor.user.first_name}"

class Accepted_Submission(models.Model):
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE)
    corrected_file = models.FileField(upload_to='submissions/')
    typeset_file = models.FileField(upload_to='submissions/')
    corrected_title = models.CharField(max_length=255, null=True, blank=True)
    corrected_abstract = models.TextField(null=True, blank=True)
    corrected_on = models.DateField(default=now)
    typeset_on = models.DateField(default=now)
    
    def __str__(self):
        return self.submission.title
    
    
class Funder(models.Model):
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE)
    detail = models.TextField()

    def __str__(self):
        return self.detail
    

class CoAuthor(models.Model):
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE,null=True, blank=True)
    name = models.CharField(max_length=255,null=True, blank=True)
    email = models.EmailField(max_length=255,null=True, blank=True)
    institution = models.CharField(max_length=255,null=True, blank=True)

    def __str__(self):
        return self.name


class Keyword(models.Model):
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE)
    keyword = models.CharField(max_length=225) 

    def __str__(self):
        return self.keyword
    
class Date(models.Model):
    journal = models.ForeignKey(Journal, on_delete=models.CASCADE, null=True, blank=True)
    due_days_to_accept_invitation = models.IntegerField(null= True,blank=True)
    due_days_to_review = models.IntegerField(null= True,blank=True)
    due_days_to_minor_revision = models.IntegerField(null= True,blank=True)
    due_days_to_major_revision = models.IntegerField(null= True,blank=True)
    due_days_to_payment = models.IntegerField(null= True,blank=True)
    due_days_to_corrections = models.IntegerField(null= True,blank=True)
    due_days_to_typeset_approval = models.IntegerField(null= True,blank=True)
    due_days_to_next_step = models.IntegerField(null= True,blank=True)

    def __str__(self):
        return self.journal.title
    
class Email(models.Model):
    from_user = models.CharField(max_length=255, null=True, blank=True)
    to_user = models.ForeignKey(User, related_name='emails_received', on_delete=models.CASCADE, null=True, blank=True)
    subject = models.CharField(max_length=255)
    content = models.TextField()
    datetime = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f'Email from {self.from_user} to {self.to_user} - {self.subject}'