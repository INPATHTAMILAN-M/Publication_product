from datetime import timedelta
from pyexpat.errors import messages
import re
from django.core.mail import EmailMessage
import logging
from django.http import  HttpResponse, HttpResponseNotAllowed, HttpResponseNotFound, JsonResponse, HttpResponseBadRequest
from django.core.paginator import Paginator, PageNotAnInteger, EmptyPage
from django.urls import reverse
from account.sms import send_sms
from account.whatsapp import send_whatsapp
from dl.models import Issue, Published_article, Volume
from .auth import *
from .forms import *
from .models import *
from account.models import Author, Editor, FeedbackType, Modes, Question
import json
from django.db.models import Q
from django.shortcuts import redirect, render, get_object_or_404
from django.contrib.auth.models import User
from django.contrib.auth.models import Group
from django.utils import timezone
from django.contrib.auth.decorators import login_required , user_passes_test
from django.utils.timezone import now 
from django.http import JsonResponse
from django.contrib.auth.models import User
from django.core.mail import send_mail
from django.views.decorators.csrf import csrf_protect
from .services import send_email
from django.db.models.functions import ExtractMonth, ExtractYear
from django.db.models import Count, Q
import datetime
from account.extractors import extract_docx, extract_pdf
from account.converter import text_to_latex
from django.core.files.storage import FileSystemStorage
from django.core.files import File
import traceback


# -----------------------------------------------------------------------------Submission Flow------------------------------------------------------------------------------------------------------------------------------------------------------------

 


#Draft menu
@login_required
@user_passes_test(is_author)

def Draftview(request):
    author = Author.objects.get(user=request.user)
    submissions = Submission.objects.filter(article_status__article_status='Draft', author=author)
    submission_statuses = check_submission_status(author)
    
    if request.method =="POST":
        action = request.POST.get('action')
        submission_id = request.POST.get('submission_id')
        if action == 'delete' and submission_id:
            submission = Submission.objects.get(id=submission_id)
            submission.delete()
            return redirect('draft')
    return render(request, 'draftview.html', {
        'submissions': submissions,
        **submission_statuses,
    })


#Submitted menu
@login_required
@user_passes_test(is_author)
def Submittedview(request):
    author = Author.objects.get(user=request.user)
    submissions = Submission.objects.filter(
        article_status__article_status__in=[
            'Submitted', 
            'Assign AE',
            'Awaiting for Reviewers',
            'Under Review',
            'Awaiting AE Recommendation',
            'Awaiting for EIC Decision',
            'Awaiting for Revision',
            'Published',
            ], author=author)
    
    paginator = Paginator(submissions, 10)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    submission_statuses = check_submission_status(author)
    
    submission_statuses = check_submission_status(author)
    
    admin_group = Group.objects.get(name='Admin Office')

# Get the first user in the 'Admin Office' group (assuming there's only one admin)
    admin_user = admin_group.user_set.first()
    
    return render(request, 'submittedview.html', {
        'submissions': submissions,
        **submission_statuses,
        'admin_email': admin_user.email if admin_user else '',
    })
    
from django.template.loader import render_to_string
#New submission start
@login_required
@user_passes_test(is_author)
def startnew(request):
    author = Author.objects.get(user=request.user)
    submission_statuses = check_submission_status(author)
    status_filter = request.GET.get('article_status', '')
    submissions = Submission.objects.filter(author=author)
    
    if status_filter:
        submissions = submissions.filter(article_status__article_status=status_filter)
    else:
        submissions = submissions.filter(article_status__article_status__in=[
            'Submitted', 
            'Assign AE',
            'Awaiting for Reviewers',
            'Under Review',
            'Awaiting AE Recommendation',
            'Awaiting for EIC Decision',
            'Awaiting for Revision',
            'Published',
        ])

    if request.headers.get('x-requested-with') == 'XMLHttpRequest':
        html = render_to_string('partials/_submission_list.html', {'submissions': submissions})
        return JsonResponse({'html': html})

    admin_group = Group.objects.get(name='Admin Office')
    admin_user = admin_group.user_set.first()

    return render(request, 'start-new-submission.html', {
        **submission_statuses,
        'submissions': submissions,
        'admin_email': admin_user.email if admin_user else '',
        'selected_status': status_filter,
    }) 

@login_required
@user_passes_test(is_author)    
def upload_additional_file(request):
    if request.method == 'POST':
        submission_id = request.POST.get('submission_id')
        submission = get_object_or_404(Submission, id=submission_id)
        additional_file = request.FILES.get('additional_file')

        if additional_file:
            submission.additional_file = additional_file
            submission.save()
            messages.success(request, 'Additional file uploaded successfully.')
        else:
            messages.error(request, 'Please select a file to upload.')

        return redirect('submitted')



#Revision menu
@login_required
@user_passes_test(is_author)
def Revisionview(request):
    author = Author.objects.get(user=request.user)
    submissions = Submission.objects.filter(article_status__article_status__in=['Minor Revision' , 'Major Revision'], author=author)
    submission_statuses = check_submission_status(author)
    return render(request,'revisionview.html', {
        'submissions': submissions,
        **submission_statuses,
    })

#Accepted menu
@login_required
@user_passes_test(is_author)
def Acceptedview(request):
    author = Author.objects.get(user=request.user)
    submissions = Submission.objects.filter(article_status__article_status__in=['Accepted'], decision__decision='Accepted',author=author)
    paginator = Paginator(submissions, 10)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    
    submission_statuses = check_submission_status(author)
    return render(request,'acceptedview.html', {
        'submissions': submissions,
        **submission_statuses,
    })
    
def correction_comments_api(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    comments = submission.correction_comments.all()
    
    comments_list = [
        {
            "id": c.id,
            "correction_commments": c.correction_commments,
            "due_date": c.due_date.strftime('%d-%b-%Y') if c.due_date else '',
        }
        for c in comments
    ]
    
    return JsonResponse({'comments': comments_list})

def upload_additional_corrected_file(request):
    if request.method == 'POST':
        comment_id = request.POST.get('comment_id')
        print(f"Comment ID: {comment_id}")  # Debugging line to check comment_id
        comment = get_object_or_404(Correction_Comments, id=comment_id)

        uploaded_file = request.FILES.get('additional_file')
        if uploaded_file:
            comment.additional_file = uploaded_file
            comment.save()
            # Return JSON success + file URL
            return JsonResponse({
                'success': True,
                'file_url': comment.additional_file.url,
            })
        else:
            return JsonResponse({'success': False, 'error': 'No file uploaded'})
    return JsonResponse({'success': False, 'error': 'Invalid request method'})

# copyright
@login_required
@user_passes_test(is_author)
def upload_copyright_form(request, submission_id):
    if request.method == 'POST':
        author = Author.objects.get(user=request.user)
        submission = get_object_or_404(Submission, id=submission_id, author=author)
        # accepted_submission = get_object_or_404(Accepted_Submission, submission=submission)
        
        # Handle file upload
        copyright_form = request.FILES.get('copyright_form')
        if copyright_form:
            submission.copyright_file = copyright_form
            submission.save()
            messages.success(request, 'Copyright form uploaded successfully!')
        else:
            messages.error(request, 'Please select a file to upload.')

    return redirect('accepted')


#Rejected menu
@login_required
@user_passes_test(is_author)
def Rejectedview(request):
    author = Author.objects.get(user=request.user)
    submissions = Submission.objects.filter(decision__decision='Rejected', author=author)
    
    paginator = Paginator(submissions, 10)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    reviewer = Author.objects.get(user=User.objects.get(username="mukesh@gmail.com"))
    submission_statuses = check_submission_status(author)
    admin_group = Group.objects.get(name='Admin Office')
    admin_user = admin_group.user_set.first()
    return render(request,'rejectedview.html', {
        'submissions': submissions,
        **submission_statuses,
        'admin_email': admin_user if admin_user else '',
        'reviewer': reviewer,
    })

#Step-1
@login_required
@user_passes_test(is_author)
def submission_step_one(request, submission_id):
    submission = None
    if submission_id and submission_id != 0:
        try:
            submission = Submission.objects.get(id=submission_id)
            if submission.article_status.article_status in ['Minor Revision', 'Major Revision']:
                new_submission = Submission(author=submission.author)
                new_submission.parent_submission_id = submission.id
                try:
                    new_submission.manuscript_id = generate_manuscript_id(submission=submission)
                except ValueError as e:
                    messages.error(request, str(e))
                    return redirect('new_submission', submission_id=submission_id)
                new_submission.article_status_id = get_object_or_404(Article_Status, article_status="Draft").id
                new_submission.save()
                request.session['submission_id'] = new_submission.id
                return redirect('new_submission', submission_id=new_submission.id)
        except Submission.DoesNotExist:
            submission_id = None

    if request.method == 'POST':
        form = SubmissionStepOneForm(request.POST, instance=submission)
        if form.is_valid():
            submission = form.save(commit=False)
            author = Author.objects.get(user=request.user)
            submission.author = author

            if not submission.manuscript_id:
                try:
                    submission.manuscript_id = generate_manuscript_id(request=request)
                except ValueError as e:
                    messages.error(request, str(e))
                    return render(request, 'begin-new-submission.html', {
                        'form': form,
                        'article_types': Article_Type.objects.all(),
                        'categories': Category.objects.all(),
                        'journals': Journal.objects.all(),
                        'submission_id': submission_id,
                        'error': str(e)
                    })
                submission.article_status_id = get_object_or_404(Article_Status, article_status="Draft").id

            submission.save()
            request.session['submission_id'] = submission.id

            #if submission_id == 0:
            #    return redirect('new_submission', submission_id=submission.id)

            if 'save_and_continue' in request.POST:
                return redirect('submission_step_two', submission_id=submission.id)
            elif submission_id == 0:
                return redirect('new_submission', submission_id=submission.id)
            else:
                return render(request, 'begin-new-submission.html', {
                    'form': form,
                    'article_types': Article_Type.objects.all(),
                    'categories': Category.objects.all(),
                    'journals': Journal.objects.all(),
                    'saved': True,
                    'submission_id': submission.id
                })
    else:
        form = SubmissionStepOneForm(instance=submission)

    return render(request, 'begin-new-submission.html', {
        'form': form,
        'article_types': Article_Type.objects.all(),
        'categories': Category.objects.all(),
        'journals': Journal.objects.all(),
        'submission_id': submission_id
    })
# *************************************
# def generate_manuscript_id(request=None, submission=None):
#     if submission:
#         journal = submission.journal
#         if not journal:
#             raise ValueError("Existing submission does not have an associated journal")
#     elif request and request.POST:
#         journal_id = request.POST.get('journal')
#         if not journal_id:
#             raise ValueError("No journal selected. Please select a journal before submitting.")
#         try:
#             journal = Journal.objects.get(id=journal_id)
#         except Journal.DoesNotExist:
#             raise ValueError(f"No Journal found with id: {journal_id}")
#     else:
#         raise ValueError("Either request with POST data or existing submission is required")

#     current_year = datetime.now().year
#     current_month = datetime.now().month
    
#     # Count how many submissions have been made to the same journal in the same year and month
#     count = Submission.objects.filter(
#         journal=journal,
#         submitted_on__year=current_year,
#         submitted_on__month=current_month
#     ).count()

#     # Increment the count by 1 for the new submission
#     count += 1

#     # Format the manuscript_id
#     manuscript_id = f"{journal.title}-{current_year}-{str(current_month).zfill(2)}-{count}"

#     return manuscript_id
def generate_manuscript_id(request=None, submission=None):
    if submission:
        journal = submission.journal
        if not journal:
            raise ValueError("Existing submission does not have an associated journal")
    elif request and request.POST:
        journal_id = request.POST.get('journal')
        if not journal_id:
            raise ValueError("No journal selected. Please select a journal before submitting.")
        try:
            journal = Journal.objects.get(id=journal_id)
        except Journal.DoesNotExist:
            raise ValueError(f"No Journal found with id: {journal_id}")
    else:
        raise ValueError("Either request with POST data or existing submission is required")

    current_year = datetime.now().year
    current_month = datetime.now().month

    # Extract the maximum manuscript number for the current year and month
    max_manuscript = Submission.objects.filter(
        journal=journal,
        submitted_on__year=current_year,
        submitted_on__month=current_month
    ).values_list('manuscript_id', flat=True).order_by('id').last()

    if max_manuscript:
        # Extract the numeric part of the manuscript_id
        match = re.match(rf"^{journal.title}-{current_year}-{str(current_month).zfill(2)}-(\d+)$", max_manuscript)
        if match:
            max_number = int(match.group(1))
            new_number = max_number + 1
        else:
            raise ValueError(f"Unexpected format in manuscript_id: {max_manuscript}")
    else:
        # No previous manuscripts, start from 1
        new_number = 1

    # Format the new manuscript_id with 'K' before the journal title
    manuscript_id = f"K{journal.title}-{current_year}-{str(current_month).zfill(2)}-{new_number}"

    return manuscript_id

    '''latest_volume = Volume.objects.filter(journal=journal).order_by('-year').first()
    if not latest_volume:
        raise ValueError(f"No volumes found for journal: {journal.title}")

    volume_year = latest_volume.year

    latest_issue = Issue.objects.filter(volume=latest_volume).order_by('-issue').first()
    if not latest_issue:
        raise ValueError(f"No issues found for the latest volume of journal: {journal.title}")

    issue_number = latest_issue.issue

    submission_count = Submission.objects.filter(
        journal=journal,
        submitted_on__year=volume_year
    ).count() + 1

    return f"{journal.title}-{volume_year}-{issue_number}-{submission_count}"'''
# *************************************************************

#Step-3
@login_required
@user_passes_test(is_author)
def keyword(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)

    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'add':
            keyword_text = request.POST.get('keyword')
            if keyword_text:
                keyword = Keyword.objects.create(submission=submission, keyword=keyword_text)
                return JsonResponse({'success': True, 'keyword': keyword.keyword, 'keyword_id': keyword.id})
            
        elif action == 'remove':
            keyword_id = request.POST.get('keyword_id')
            Keyword.objects.filter(id=keyword_id, submission=submission).delete()
            return JsonResponse({'success': True, 'keyword_id': keyword_id})
        
        elif action == 'save_continue':
            request.session['submission_id'] = submission.id
            return redirect('step4', submission_id=submission.id)

    context = {
        'submission': submission,
        'keywords': Keyword.objects.filter(submission=submission),
    }
    return render(request, 'step-3.html', context)

#Step-4
@login_required
@user_passes_test(is_author)
def add_authors_institutions(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    coauthor_form = CoAuthorForm()

    if request.method == "POST":
        if 'save_and_continue' in request.POST:
            request.session['submission_id'] = submission.id
            return redirect('step5', submission_id=submission.id)
            
    coauthors = CoAuthor.objects.filter(submission=submission)
    context = {
        'coauthor_form': coauthor_form,
        'submission': submission,
        'current_author': Author.objects.get(user=request.user),
        'coauthors': coauthors,
    }
    return render(request, 'step-4.html', context)

def add_coauthor_ajax(request, submission_id):
    if request.method == 'POST':
        form = CoAuthorForm(request.POST, submission=submission_id)
        if form.is_valid():
            coauthor = form.save(commit=False)
            coauthor.submission_id = submission_id
            coauthor.save()
            response = {
                'id': coauthor.id,
                'name': coauthor.name,
                'email': coauthor.email,
                'institution': coauthor.institution,
            }
            return JsonResponse(response, status=201)
        else:
            return JsonResponse({'errors': form.errors}, status=400)

def remove_coauthor_ajax(request, submission_id):
    if request.method == 'POST':
        coauthor_id = request.POST.get('coauthor_id')
        CoAuthor.objects.filter(id=coauthor_id, submission_id=submission_id).delete()
        return JsonResponse({'id': coauthor_id}, status=200)

#Step-5
@login_required
@user_passes_test(is_author)
def step5(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)

    try:
        funder = Funder.objects.get(submission=submission)
    except Funder.DoesNotExist:
        funder = None

    if request.method == "POST":
        form = SubmissionForm(request.POST, instance=submission)
        funder_form = FunderForm(request.POST, instance=funder)
        if form.is_valid() and (not submission.is_funded or (submission.is_funded and funder_form.is_valid())):
            form.save()

            if submission.is_funded:
                funder = funder_form.save(commit=False)
                funder.submission = submission
                funder.save()
            action = request.POST.get('action')
            if action == 'save_continue':
                return redirect('step6_review_submit', submission_id=submission.id)
            else:
                return render(request, 'step-5.html', {
                    'form': form,
                    'funder_form': funder_form,
                    'submission': submission,
                })
    else:
        form = SubmissionForm(instance=submission)
        funder_form = FunderForm(instance=funder)

    return render(request, 'step-5.html', {
        'form': form,
        'funder_form': funder_form,
        'submission': submission,
        'specialization':Specialization.objects.all(),
    })

#Step-6
@login_required
@user_passes_test(is_author)
def step6_review_submit(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    submission_files = Submission_Files.objects.filter(submission=submission)
    keywords = Keyword.objects.filter(submission=submission)
    coauthors = CoAuthor.objects.filter(submission=submission)
    try:
        funder = Funder.objects.get(submission=submission)
    except Funder.DoesNotExist:
        funder = None

    empty_fields = []

    # Check each field and add to empty_fields if necessary
    if not submission.title:
        empty_fields.append('Title')
    if not submission.abstract:
        empty_fields.append('Abstract')
    if not submission.cover_letter:
        empty_fields.append('Cover Letter')
    # if submission.no_of_figures is None:
    #     empty_fields.append('Number of Figures')
    # if submission.no_of_tables is None:
    #     empty_fields.append('Number of Tables')
    # if not submission.no_of_words:
    #     empty_fields.append('Number of Words')
    if not submission.specialization:
        empty_fields.append('Specialization')
    if not submission.coi_describe and submission.conflict_of_interest ==1  :
        empty_fields.append('Conflict of Interest Description')
    if not submission.category:
        empty_fields.append('Category')
    if not submission.journal:
        empty_fields.append('Journal')
    if not submission.article_type:
        empty_fields.append('Article Type')
    if not submission_files.exists():
        empty_fields.append('Submission Files')
    if not keywords.exists():
        empty_fields.append('Keywords')
    # if not coauthors.exists():
    #     empty_fields.append('Co-Authors')
    if submission.is_funded == 1 and funder == None:
        empty_fields.append('Funder information')
    if submission.is_funded == 1 and submission.is_funded == 0:
        empty_fields.append(' is Funded')

    form = ReviewSubmitForm(initial={
            'title': submission.title,
            'abstract': submission.abstract,
            'cover_letter': submission.cover_letter,
            'is_funded': submission.is_funded,
            'figures': submission.no_of_figures,
            'tables': submission.no_of_tables,
            'specialization': submission.specialization,
            'words': submission.no_of_words,
            'is_submitted_already': submission.is_submitted_already,
            'acknowledgement_1': submission.acknowledgement_1,
            'acknowledgement_2': submission.acknowledgement_2,
            'acknowledgement_3': submission.acknowledgement_3,
            'conflict_of_interest': submission.conflict_of_interest,
            'coi_describe': submission.coi_describe if submission.conflict_of_interest == 1 else 'No',
            
        })

    context = {
        'form': form,
        'submission': submission,
        'category': submission.category,
        'journal': submission.journal,
        'Article_Type': submission.article_type,
        'submission_files': submission_files,
        'keywords': keywords,
        'coauthors': coauthors,
        'funder': funder,
        'empty_fields': empty_fields,
        'specialization': submission.specialization,

    }
    return render(request, 'step-6.html', context)


#**********************************************************************************************************************

from django.core.files.base import ContentFile
from django.core.files.storage import default_storage
from django.shortcuts import get_object_or_404, redirect, render
from django.http import FileResponse, HttpResponseBadRequest
from django.conf import settings
from django.contrib.auth.decorators import login_required, user_passes_test
from PyPDF2 import PdfWriter, PdfReader
from PIL import Image
import os
import logging
from .models import Submission, Submission_Files, Article_Status
from .forms import SubmissionFileForm
import subprocess
import platform

logger = logging.getLogger(__name__)

# Step 2: Handle file uploads
@login_required
@user_passes_test(is_author)
def submission_step_two(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    submission_files = Submission_Files.objects.filter(submission=submission)

    if request.method == 'POST':
        if 'remove_file' in request.POST:
            file_id = request.POST.get('file_id')
            file_to_remove = get_object_or_404(Submission_Files, id=file_id)
            file_to_remove.delete()
            return redirect('submission_step_two', submission_id=submission.id)
        else:
            form = SubmissionFileForm(request.POST, request.FILES, initial={'submission': submission})
            if form.is_valid():
                files = request.FILES.getlist('file')
                file_category_ids = request.POST.getlist('file_category')  # Fetch the list of category IDs

                if len(files) != len(file_category_ids):
                    form.add_error(None, "The number of files must match the number of categories.")
                else:
                    for file, category_id in zip(files, file_category_ids):
                        file_size = file.size
                        # Fetch the File_Category instance using the category_id
                        file_category = get_object_or_404(File_Category, id=category_id)
                       
                        Submission_Files.objects.create(
                            submission=submission,
                            file=file,
                            file_category=file_category,  # Use the instance, not the ID
                            file_size=file_size
                        )

                    return redirect('submission_step_two', submission_id=submission.id)
    else:
        form = SubmissionFileForm(initial={'submission': submission})

    return render(request, 'step-2.html', {'form': form, 'submission': submission, 'submission_files': submission_files})
    
# View proof in Step 6
def view_proof(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    submission_files = Submission_Files.objects.filter(submission=submission)

    logger.debug(f"Attempting to process {len(submission_files)} files for submission ID: {submission_id}")

    try:
        merged_pdf_path = _process_files(submission_files)

        # Save the merged PDF to the final_file field of the Submission model
        submission.final_file.save(f'merged_submission_{submission.id}.pdf', open(merged_pdf_path, 'rb'))
        submission.save()

        
        os.remove(merged_pdf_path)

        # Return the merged PDF as a response
        response = FileResponse(open(submission.final_file.path, 'rb'), content_type='application/pdf')
        response['Content-Disposition'] = f'inline; filename="merged_submission_{submission.id}.pdf"'

        return response

    except Exception as e:
        logger.error(f"Error generating merged PDF for submission ID {submission_id}: {str(e)}")
        return HttpResponseBadRequest("An error occurred while generating the merged PDF. Please try again later.")

# File conversion and merging
def convert_docx_to_pdf_with_libreoffice(input_path, output_path):
    # Determine the LibreOffice path based on the operating system
    if platform.system() == 'Windows':
        libreoffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
    else:  # For Linux (e.g., Ubuntu) or macOS
        libreoffice_path = '/usr/bin/soffice'
    
    if not os.path.exists(libreoffice_path):
        raise FileNotFoundError(f"LibreOffice executable not found at: {libreoffice_path}")
    
    result = subprocess.run([libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(output_path), input_path])
    if result.returncode != 0:
        raise Exception(f"Error converting DOCX to PDF: {result.returncode}")

def _process_files(files):
    output_pdf_path = os.path.join(settings.MEDIA_ROOT, 'merged_submission.pdf')
    pdf_writer = PdfWriter()

    for file in files:
        try:
            file_extension = os.path.splitext(file.file.name)[1].lower()
            temp_path = default_storage.path(file.file.name)

            logger.debug(f"Processing file: {temp_path}")

            if not os.path.exists(temp_path):
                logger.error(f"File not found at path: {temp_path}")
                continue

            if file_extension == '.docx':
                temp_pdf_path = os.path.join(settings.MEDIA_ROOT, f"{os.path.splitext(file.file.name)[0]}.pdf")

                try:
                    convert_docx_to_pdf_with_libreoffice(temp_path, temp_pdf_path)
                except Exception as e:
                    logger.error(f"Error converting DOCX to PDF: {e}")
                    continue

                if os.path.exists(temp_pdf_path):
                    with open(temp_pdf_path, 'rb') as pdf_file:
                        pdf_reader = PdfReader(pdf_file)
                        for page in pdf_reader.pages:
                            pdf_writer.add_page(page)
                    default_storage.delete(temp_pdf_path)
                else:
                    logger.error(f"Converted PDF not found: {temp_pdf_path}")

            elif file_extension in ['.jpg', '.jpeg', '.png', '.gif']:
                temp_pdf_path = os.path.join(settings.MEDIA_ROOT, f"{os.path.splitext(os.path.basename(temp_path))[0]}.pdf")

                try:
                    image = Image.open(temp_path)
                    image.convert('RGB').save(temp_pdf_path)

                    if os.path.exists(temp_pdf_path):
                        with open(temp_pdf_path, 'rb') as pdf_file:
                            pdf_reader = PdfReader(pdf_file)
                            for page in pdf_reader.pages:
                                pdf_writer.add_page(page)
                        default_storage.delete(temp_pdf_path)
                    else:
                        logger.error(f"Converted PDF not found: {temp_pdf_path}")
                except Exception as e:
                    logger.error(f"Error converting image to PDF: {e}")

            elif file_extension == '.pdf':
                try:
                    if os.path.exists(temp_path):
                        with open(temp_path, 'rb') as pdf_file:
                            pdf_reader = PdfReader(pdf_file)
                            for page in pdf_reader.pages:
                                pdf_writer.add_page(page)
                    else:
                        logger.error(f"PDF file not found: {temp_path}")
                except Exception as e:
                    logger.error(f"Error processing PDF file: {e}")

        except Exception as e:
            logger.error(f"Unexpected error processing file {file.file.name}: {e}")

    try:
        with open(output_pdf_path, 'wb') as output_pdf_file:
            pdf_writer.write(output_pdf_file)
        logger.debug(f"Merged PDF saved at: {output_pdf_path}")
    except Exception as e:
        logger.error(f"Error saving merged PDF: {e}")
        raise Exception("An error occurred while generating the merged PDF. Please try again later.")

    return output_pdf_path

#Submit in Step-6
def submission_step_six(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    article_status_submitted = get_object_or_404(Article_Status, article_status='submitted')
    
    if request.method == 'POST':
        if 'action' in request.POST:
            if request.POST['action'] == 'submit':
                submission.article_status_id = article_status_submitted.id
                submission.save()
                if Modes.objects.filter(name="Email",is_active=True):
                    send_email(
                        to_email=submission.author.user.email,
                        subject='Paper Status Update - ICERCS 2024',
                        template_name='email_templates/submitted_author.html',
                        user=submission.author.user,
                        context={'submission': submission }
                    )
                if Modes.objects.filter(name="Whatsapp",is_active=True):
                    send_whatsapp(
                        user=submission.author,
                        template_name='email_templates/submitted_author.html',
                        context={'submission': submission }
                    )
                if Modes.objects.filter(name="SMS",is_active=True):
                    send_sms(
                        user=submission.author.user,
                        template_name='email_templates/submitted_author.html',
                        context={'submission': submission }
                    )
                admin =User.objects.get(
                    groups__name='Admin Office'
                )
                
                send_email(
                    to_email=admin.email,
                    subject='Manuscript Submitted',
                    template_name='email_templates/submitted_admin.html',
                    user=admin,
                    context={'submission': submission , 'admin':admin}
                )
                
            # Redirect to a confirmation page or wherever needed
            return redirect('submitted')  # Replace 'confirmation_page' with your actual URL name
    return render(request, 'step-6.html', {'submission': submission})

#*****************************************************************************
#Author dashboard avail checker
def check_submission_status(author):
    return {
        'has_draft': Submission.objects.filter(article_status__article_status='Draft', author=author).exists(),
        'has_submitted': Submission.objects.filter(article_status__article_status__in=[
            'Submitted', 
            'Assign AE',
            'Awaiting for Reviewers',
            'Under Review',
            'Awaiting AE Recommendation',
            'Awaiting for EIC Decision',
            'Awaiting for Revision',
            'Published',
            ], author=author).exists(),
        'has_revision': Submission.objects.filter(article_status__article_status__in=['Minor Revision', 'Major Revision'], author=author).exists(),
        'has_accepted': Submission.objects.filter(decision__decision='Accepted', author=author).exists(),
        'has_rejected': Submission.objects.filter(decision__decision='Rejected', author=author).exists(),
    }





# --------------------------------------------------------------------------------Admin Office-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

from django.core.paginator import Paginator

@login_required
@user_passes_test(is_admin_office)

def admin_office_view(request):
    submission = Submission.objects.filter(article_status__article_status='Submitted')
    editors = Editor.objects.all()
    journals = Journal.objects.all()
    
    # Set up pagination
    paginator = Paginator(submission, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    skip_ae = Modes.objects.filter(name='Skip AE', is_active=True).exists()
    return render(request, 'admin_office.html', {'submissions': submissions, 'editors': editors, 'journals': journals, 'skip_ae': skip_ae})

def move_to_eic(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    eic_status = get_object_or_404(Article_Status, article_status="Awaiting for EIC Review")
    submission.article_status = eic_status
    submission.save()
    return redirect('admin_office')

from django.contrib import messages

# palg report
def upload_plag_report(request):
    if request.method == 'POST':
        manuscript_id = request.POST.get('manuscript_id')
        plag_report = request.FILES.get('plag_report')

        submission = get_object_or_404(Submission, id=manuscript_id)
        submission.plag_report = plag_report
        submission.save()

        messages.success(request, 'Plagiarism report uploaded successfully!')
        return redirect('admin_office')

def remove_plag_report(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    
    # Clear the plagiarism report
    submission.plag_report = None
    submission.save()

    messages.success(request, 'Plagiarism report removed successfully!')
    return redirect('admin_office')

from datetime import datetime

@login_required
@user_passes_test(is_admin_office)

def assign_ae(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    
    if request.method == 'POST':
        editor_id = request.POST.get('editor_id')
        
        # Validate and parse the date
        date = datetime.now().date()

        editor = get_object_or_404(User, editor__id=editor_id)

        # Check if an AE_Assignment already exists for this submission and editor
        assignment, created = AE_Assignment.objects.get_or_create(
            user=editor,
            submission=submission,
            defaults={
                'date': date,
                'ae_recommendation': 'A',  # Default or modify as needed
                'ae_comments': ''  # Modify as needed
            }
        )

        if not created:
            # Optionally update the existing assignment here if needed
            assignment.date = date
            assignment.save()
        submission.article_status = Article_Status.objects.get(article_status='AE Assigned') 
        submission.decissioned_on = timezone.now()
        submission.decision_by = request.user
        submission.save()
        if Modes.objects.filter(name="Email",is_active=True):
            send_email(
                        to_email=editor.email,
                        subject='Manuscript Assigned',
                        template_name='email_templates/ae_assigned.html',
                        user=editor.user,
                        context={
                            'submission': submission,
                            'editor': editor,
                        }
                    )
        if Modes.objects.filter(name="Whatsapp",is_active=True):
            send_whatsapp(
                template_name='email_templates/ae_assigned.html',
                user=editor,
                context={
                    'submission': submission,
                    'editor': editor,
                }
            )
        if Modes.objects.filter(name="SMS",is_active=True):
            send_sms(
                template_name='email_templates/ae_assigned.html',
                user=editor,
                context={
                    'submission': submission,
                    'editor': editor,
                }
            )
        return redirect('admin_office')    # Adjust this to your actual redirect path

    else:
        # Fetch editors assigned to this journal who are part of the 'AE' group
        journal = submission.journal
        journal_editors = Journal_Editor_Assignment.objects.filter(journal=journal).select_related('editor')
        ae_group = Group.objects.get(name='AE')
        ae_editors = journal_editors.filter(editor__user__groups=ae_group)

        context = {
            'submission': submission,
            'ae_editors': ae_editors,
            'date': timezone.now().date(),  # Pass the current date to the form
        }
        return render(request, 'assign_ae_modal.html', context)

@login_required

def reject_manuscript(request, manuscript_id):
    if request.method == 'POST':
        try:
            submission = Submission.objects.get(id=manuscript_id)
            data = json.loads(request.body)
            submission.admin_comments = data.get('admin_comments', '')
            submission.article_status = Article_Status.objects.get(article_status='Rejected')
            submission.decision=Decision.objects.get(decision='Rejected')
            submission.decissioned_on = timezone.now()
            submission.decision_by = request.user
            submission.save()
            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                        to_email=submission.author.user.email,
                        subject='Manuscript Rejected',
                        template_name='email_templates/decision_rejected.html',
                        user=submission.author.user,
                        context={'submission': submission }
                    )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=submission.author,
                    template_name='email_templates/decision_rejected.html',
                    context={'submission': submission }
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=submission.author,
                    template_name='email_templates/decision_rejected.html',
                    context={'submission': submission }
                )
            return JsonResponse({'success': True})
        except Submission.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'Submission not found'})
        except json.JSONDecodeError:
            return JsonResponse({'success': False, 'error': 'Invalid JSON'})
        except Exception as e:
            return JsonResponse({'success': False, 'error': str(e)})
    return JsonResponse({'success': False, 'error': 'Invalid request method'})

# ------Rejected Manuscripts-------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_admin_office)
def Manuscripts_rejection(request):
    rejected_status = Article_Status.objects.get(article_status='Rejected')
    submissions = Submission.objects.filter(article_status=rejected_status).order_by('-id')
     # Set up pagination
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    return render(request, 'manuscripts_rejection.html', {
        'submissions': submissions, 
        'feedback_types': FeedbackType.objects.all(),
        'all_questions': Question.objects.all()
        })

# ------Accepted Manuscripts----------------------------------------------------------------------------------------------------------------------------------------------------------------

from django.db.models.functions import Coalesce
from django.db.models import F
from django.db.models import Case, When, Value, IntegerField


@login_required
@user_passes_test(is_admin_office)
def Manuscripts_acceptance(request):
    # submissions = Submission.objects.filter(article_status__article_status='Awaiting for Correction after Acceptance').order_by('submitted_on')
    # submissions += Submission.objects.filter(article_status__article_status='Payment Done').order_by('payment_date')
    # submissions += Submission.objects.filter(article_status__article_status='Accepted').order_by('submitted_on')
    # submissions = Submission.objects.filter(
    #     article_status__article_status__in=['Accepted', 'Payment Done', 'Awaiting for Correction after Acceptance']
    # ).annotate(
    #     sort_date=Coalesce('payment_date', 'submitted_on')
    # ).order_by('sort_date')


    submissions = Submission.objects.filter(
        article_status__article_status__in=[
            'Awaiting for Correction after Acceptance',
            'Payment Done',
            'Accepted'
        ]
    ).annotate(
        custom_order=Case(
            When(article_status__article_status='Awaiting for Correction after Acceptance', then=Value(1)),
            When(article_status__article_status='Payment Done', then=Value(2)),
            When(article_status__article_status='Accepted', then=Value(3)),
            output_field=IntegerField()
        ),
        sort_date=Case(
            When(article_status__article_status='Payment Done', then='payment_date'),
            default='submitted_on'
        )
    ).order_by('custom_order', 'sort_date')

     # Set up pagination
    paginator = Paginator(submissions, 20)  
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    Context={
        'submissions': submissions,
        'feedback_types': FeedbackType.objects.all(),
        'all_questions': Question.objects.all(),
    }
    return render(request, 'manuscripts_acceptance.html', Context)


logger = logging.getLogger(__name__)

def send_correction_report(request):
    if request.method == 'POST':
        submission_id = request.POST.get('submission_id')
        correction_comments = request.POST.get('correction_comments')
        try:
            submission = Submission.objects.get(id=submission_id)
            journal = submission.journal
            date_obj = Date.objects.get(journal=journal)
            due_days = date_obj.due_days_to_corrections
            submission.correction_due_date = timezone.now().date() + timedelta(days=due_days)

            # Retrieve all matching Correction_Comments objects
            correction_comment_objs = Correction_Comments.objects.filter(submission=submission)

            if correction_comment_objs.exists():
                # Update the first matching Correction_Comments object
                correction_comment_obj = correction_comment_objs.first()
                correction_comment_obj.correction_commments = correction_comments
                correction_comment_obj.due_date = timezone.now().date() + timedelta(days=due_days)
                correction_comment_obj.save()
            else:
                # Create a new Correction_Comments object
                Correction_Comments.objects.create(
                    submission=submission,
                    correction_commments=correction_comments,
                    due_date=timezone.now().date() + timedelta(days=due_days)
                )
            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                        to_email=submission.author.user.email,
                        subject='Corrections in Manuscript ',
                        template_name='email_templates/corrections_to_author.html',
                        user=submission.author.user,
                        context={'submission': submission }
                    )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=submission.author,
                    template_name='email_templates/corrections_to_author.html',
                    context={'submission': submission }
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=submission.author,
                    template_name='email_templates/corrections_to_author.html',
                    context={'submission': submission }
                )

            submission.article_status = Article_Status.objects.get(article_status='Awaiting for Correction after Acceptance')  
            submission.decissioned_on = timezone.now()
            submission.decision_by = request.user
            submission.save()

            return JsonResponse({'message': 'Correction report sent successfully!'})
        except Submission.DoesNotExist:
            return JsonResponse({'error': 'Submission not found.'}, status=404)
        except Exception as e:
            logger.error(f"An error occurred: {e}")
            return JsonResponse({'error': str(e)}, status=500)

def get_correction_file(request):
    submission_id = request.GET.get('submission_id')
    comment = Correction_Comments.objects.filter(submission_id=submission_id).last()

    if comment and comment.additional_file:
        return JsonResponse({'file_url': comment.additional_file.url})
    else:
        return JsonResponse({'file_url': None})
    
def get_submission_details(request):
    if request.method == 'GET':
        submission_id = request.GET.get('submission_id')
        submission = Submission.objects.get(id=submission_id)
        data = {
            'title': submission.title,
            'abstract': submission.abstract,
        }
        return JsonResponse(data)
    return JsonResponse({'message': 'Invalid request'}, status=400)



logger = logging.getLogger(__name__)
@login_required
@user_passes_test(is_admin_office)
def upload_corrected_file(request):
    if request.method == 'POST':
        form = CorrectedFileForm(request.POST, request.FILES)
        if form.is_valid():
            submission_id = form.cleaned_data['submission_id']
            submission = Submission.objects.get(id=submission_id)
            corrected_file = form.cleaned_data['corrected_file']
            corrected_title = form.cleaned_data['corrected_title']
            corrected_abstract = form.cleaned_data['corrected_abstract']

            # Ensure the corrected_on field is set when creating the object
            accepted_submission, created = Accepted_Submission.objects.get_or_create(
                submission=submission,
                defaults={'corrected_on': now().date(), 'typeset_on': now().date()}  # Set defaults for required fields
            )

            accepted_submission.corrected_file = corrected_file
            accepted_submission.corrected_title = corrected_title
            accepted_submission.corrected_abstract = corrected_abstract

            # If the object already existed, set the corrected_on field
            if not created:
                accepted_submission.corrected_on = now().date()
            
            accepted_submission.save()
            
            submission.article_status = Article_Status.objects.get(article_status='Corrected')
            
            submission.save()
            return JsonResponse({'message': 'Corrected file uploaded successfully!'})
        else:
            return JsonResponse({'message': 'Invalid form data'}, status=400)

    return JsonResponse({'message': 'Invalid request'}, status=400)
# --------manuscripts under review------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_admin_office)
def manuscripts_review(request):
    submissions = Submission.objects.filter(article_status__article_status='Under Review').order_by('-id')
     # Set up pagination
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    return render(request, 'manuscripts_review.html', {'submissions': submissions})


# ------Waiting Manusciripts--------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_admin_office)
def Manuscripts_revision(request):
    submissions = Submission.objects.filter(article_status__article_status='Awaiting for revision')
     # Set up pagination
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    return render(request, 'manuscripts_revision.html', {'submissions': submissions})
@login_required
@user_passes_test(is_admin_office)
def submission_details(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    return render(request, 'submission_details.html', {'submission': submission})

@login_required
@user_passes_test(is_admin_office)
@csrf_protect 
def reject_submission(request):
    if request.method == 'POST':        
        submission = get_object_or_404(Submission, id=request.POST.get('submission_id'))
        submission.admin_comments = request.POST.get('admin_comments')
        submission.save()
        
        return redirect('manuscripts_revision')
    
    return HttpResponse(status=405)  # Method not allowed
# ------------Manuscripts overdue---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_admin_office)
def Manuscripts_revision_overdue(request):
    current_date = timezone.now().date()
    submissions = Submission.objects.filter(revision_due_date__lt=current_date)
    return render(request, 'manuscripts_revision_overdue.html', {'submissions': submissions})

# -----Type Setting & Proof Reading -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_admin_office)
def Setting_proof(request):
    submissions = Submission.objects.filter(
        article_status__article_status__in=['Awaiting for Proof Read', 'Awaiting Changes', 'Corrected', 'Ready To Publish']
    ) 
     # Set up pagination
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    return render(request, 'setting_proof.html', {'submissions': submissions})

# def upload_typeset_document(request):
#     if request.method == 'POST' and request.headers.get('x-requested-with') == 'XMLHttpRequest':
#         submission_id = request.POST.get('submission_id')
#         typeset_file = request.FILES.get('typeset_file')

#         if not submission_id or not typeset_file:
#             return JsonResponse({'success': False, 'error': 'Missing submission ID or file'})

#         submission = get_object_or_404(Submission, manuscript_id=submission_id)
#         accepted_submission, created = Accepted_Submission.objects.get_or_create(submission=submission)
#         accepted_submission.typeset_file = typeset_file
#         accepted_submission.typeset_on = now()
#         accepted_submission.save()
#         submission.article_status = Article_Status.objects.get(article_status='Awaiting for Proof Read')
#         journal = submission.journal  
#         date_obj = Date.objects.get(journal=journal)
#         due_days = date_obj.due_days_to_typeset_approval
#         submission.typeset_due_date=timezone.now().date() + timedelta(days=due_days)
#         submission.save()

#         # Send email to the author
#         author_email = submission.author.email  # Assumes Submission model has an author field with an email attribute
#         subject = 'Your Manuscript Type Set Document'
#         message = 'Please check the attached type set document for your manuscript.'

#         try:
#             email = EmailMessage(
#                 subject=subject,
#                 body=message,
#                 from_email=settings.DEFAULT_FROM_EMAIL,
#                 to=[author_email]
#             )
#             email.attach(typeset_file.name, typeset_file.read(), typeset_file.content_type)
#             email.send()
            
#         except Exception as e:
#             return JsonResponse({'success': False, 'error': str(e)})

#         return JsonResponse({'success': True})

#     return JsonResponse({'success': False, 'error': 'Invalid request'})


def upload_typeset_document(request):
    if request.method == 'POST' and request.headers.get('x-requested-with') == 'XMLHttpRequest':
        submission_id = request.POST.get('submission_id')
        typeset_file = request.FILES.get('typeset_file')

        if not submission_id or not typeset_file:
            return JsonResponse({'success': False, 'error': 'Missing submission ID or file'})

        try:
            submission = get_object_or_404(Submission, manuscript_id=submission_id)
            accepted_submission, created = Accepted_Submission.objects.get_or_create(submission=submission)

            # Save uploaded file temporarily
            fs = FileSystemStorage()
            filename = fs.save(typeset_file.name, typeset_file)
            filepath = fs.path(filename)

            # Extract text based on file extension
            if filename.endswith('.docx'):
                text = extract_docx(filepath)
            elif filename.endswith('.pdf'):
                text = extract_pdf(filepath)
            else:
                with open(filepath, 'r', encoding='utf-8') as f:
                    text = f.read()

            # Convert text to LaTeX code
            latex_code = text_to_latex(text.strip())
            print("Generated LaTeX Code:", latex_code)  # Debug print

            # Prepare LaTeX output directory and files
            output_dir = 'latex_output'
            os.makedirs(output_dir, exist_ok=True)
            tex_path = os.path.join(output_dir, 'output.tex')

            with open(tex_path, 'w', encoding='utf-8') as f:
                f.write(latex_code)

            # Compile LaTeX to PDF
            result = subprocess.run(
                ['xelatex', '-interaction=nonstopmode', '-output-directory', output_dir, tex_path],
                capture_output=True, text=True, encoding='utf-8', timeout=60
            )

            pdf_path = os.path.join(output_dir, 'output.pdf')
            if not os.path.exists(pdf_path):
                error_msg = result.stderr or "PDF generation failed"
                # Cleanup
                os.remove(filepath)
                if os.path.exists(tex_path):
                    os.remove(tex_path)
                return JsonResponse({'success': False, 'error': error_msg})

            # Save the generated PDF in the model
            with open(pdf_path, 'rb') as pdf_file:
                django_file = File(pdf_file)
                accepted_submission.typeset_file.save('converted.pdf', django_file, save=True)

            accepted_submission.typeset_on = now()
            accepted_submission.save()

            # Update submission status and due date
            submission.article_status = Article_Status.objects.get(article_status='Awaiting for Proof Read')
            journal = submission.journal
            date_obj = Date.objects.get(journal=journal)
            due_days = date_obj.due_days_to_typeset_approval
            submission.typeset_due_date = now().date() + timedelta(days=due_days)
            submission.save()

            # Prepare email details
            author_email = submission.author.user.email
            subject = 'Your Manuscript Type Set Document'
            message = 'Please check the attached type set document for your manuscript.'

            try:
                email = EmailMessage(
                    subject=subject,
                    body=message,
                    from_email=settings.DEFAULT_FROM_EMAIL,
                    to=[author_email]
                )
                with open(pdf_path, 'rb') as pdf_file:
                    email.attach('converted.pdf', pdf_file.read(), 'application/pdf')

                email.send(fail_silently=False)
                print("Email sent successfully")

            except Exception as e:
                # Cleanup temp files
                os.remove(filepath)
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                if os.path.exists(tex_path):
                    os.remove(tex_path)
                print("Error sending email:", str(e))
                return JsonResponse({'success': False, 'error': f"Email sending failed: {str(e)}"})

            # Cleanup temp files after success
            os.remove(filepath)
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            if os.path.exists(tex_path):
                os.remove(tex_path)

            return JsonResponse({'success': True})

        except Exception as e:
            # Catch any other errors in processing
            tb = traceback.format_exc()
            print("Error in upload_typeset_document:", tb)
            return JsonResponse({'success': False, 'error': f"Internal error: {str(e)}"})

    return JsonResponse({'success': False, 'error': 'Invalid request'})

def view_typeset_document(request, submission_id):
    accepted_submission = get_object_or_404(Accepted_Submission, submission__id=submission_id)
    return FileResponse(accepted_submission.typeset_file.open(), content_type='application/pdf')

def mark_proof_read_done(request, submission_id):
    submission = get_object_or_404(Submission, id=submission_id)
    if request.method == 'POST':
        submission.article_status = Article_Status.objects.get(article_status='Proof Read Done')
        submission.save()
    return redirect('setting_proof') 

# ---------------------------------------------------History of Manuscripts____________________________________________________________________________
@login_required
@user_passes_test(is_admin_office)
def history(request):
    # Get all unique article statuses
    all_statuses = Submission.objects.exclude(article_status__article_status='Draft').values_list('article_status__article_status', flat=True).distinct()

    # Get the search term from the request
    search_term = request.GET.get('search', '')
    
    # Filter submissions based on selected status and search term
    selected_status = request.GET.get('status', None)
    
    # Build the query
    submissions = Submission.objects.filter(
        article_status__article_status__in=[
            'AE Assigned', 'Awaiting for Reviewers', 'Published', 'Rejected',
            'Awaiting for Proof Read', 'Awaiting for Revision', 'In Admin Processing',
            'Awaiting for Correction after Acceptance', 'Payment Done',
            'Awaiting for EIC Decision', 'Corrected', 'Awaiting Changes',
            'Ready to Publish', 'Type Set Done', 'Under Review', 'Submitted',
            'Proof Read Done', 'Awaiting AE Recommendation', 'Accepted'
        ]
    ).order_by('-id')
    
    if selected_status:
        submissions = submissions.filter(article_status__article_status=selected_status).order_by('-id')

    if search_term:
        submissions = submissions.filter(title__icontains=search_term).order_by('-id')

    paginator = Paginator(submissions, 20)  # Show 20 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)

    return render(request, 'history.html', {
        'submissions': submissions,
        'all_statuses': all_statuses,
        'selected_status': selected_status,
        'search_term': search_term,
    })

# ---------------------------------------------------------------Editor In Chief--------------------------------------------------------------------------------------------------------------
@login_required 
@user_passes_test(is_eic)

def Editor_in_Chief(request):
    # Get the logged-in user's journal
    user = request.user
    try:
        editor = Editor.objects.get(user=user)
        journals = Journal_Editor_Assignment.objects.filter(editor=editor).values_list('journal', flat=True)
    except Editor.DoesNotExist:
        journals = []

    # Filter submissions by article status and the logged-in user's journal
    submissions = Submission.objects.filter(article_status__article_status='Awaiting for EIC Decision', journal__in=journals)
    decisions = Decision.objects.all()
    ae_assignments = {submission.id: AE_Assignment.objects.filter(submission=submission).first() for submission in submissions}

    # Set up pagination
    paginator = Paginator(submissions, 20)  # Show 20 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)

    ae_assignment = AE_Assignment.objects.select_related('user').filter(submission__in=submissions)
    skip_ae = Modes.objects.filter(name='Skip AE', is_active=True).exists()
    return render(request, 'editor_in_chief.html', {'submissions': submissions, 'ae_assignments': ae_assignments, 'decisions': decisions, 'ae_assignment': ae_assignment, 'skip_ae': skip_ae})

def eic_review_manuscripts(request):
    # Submissions with status "Awaiting for EIC Review"
    submissions_eic = Submission.objects.filter(article_status__article_status='Awaiting for EIC Review')

    # Submissions with invitations sent by request.user and status "Awaiting for Reviewers"
    reviewer_invitations = Reviewer_Invitation.objects.filter(
        invite_by=request.user,
        submission__article_status__article_status='Awaiting for Reviewers'
    ).select_related('submission')

    # Extract the related submissions from reviewer invitations
    submissions_reviewers = Submission.objects.filter(id__in=[inv.submission.id for inv in reviewer_invitations])

    # Combine both querysets (remove duplicates if any)
    combined_submissions = submissions_eic.union(submissions_reviewers)

    # Pagination
    paginator = Paginator(combined_submissions, 20)
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)

    specializations = Specialization.objects.all()

    context = {
        'submissions': submissions,
        'specializations': specializations,
    }
    return render(request, 'eic_review_manuscripts.html', context)

@login_required
@user_passes_test(is_eic)
def save_decision(request):
    if request.method == 'POST':
        submission_id = request.POST.get('submission_id')
        decision_id = request.POST.get('decision')
        comments = request.POST.get('comments', '')
        include_ae_comments = request.POST.get('include_ae_comments', True)
        submission = Submission.objects.get(id=submission_id)
        decision = Decision.objects.get(id=decision_id)

        if include_ae_comments:
            ae_assignment = AE_Assignment.objects.filter(submission=submission).first()
            if ae_assignment:
                ae_comments = ae_assignment.ae_comments.strip()  
                if not comments.endswith(ae_comments):
                    # Append AE comments with proper format
                    combined_comments = f"{comments.strip()}\n\nAE Comments:\n{ae_comments}"
                else:
                    combined_comments = comments.strip()  
            else:
                combined_comments = comments.strip()  
        else:
            combined_comments = comments.strip()  

        submission.admin_comments = combined_comments

        submission.decision = decision
        submission.is_decissioned=True
        decision=Decision.objects.get(id=decision_id)
        submission.decision_by = request.user
        if decision.decision == 'Minor Revision' or decision.decision == 'Major Revision':
            submission.article_status=Article_Status.objects.get(article_status='Awaiting for revision')

            if decision.decision == 'Minor Revision':
                journal = submission.journal  
                date_obj = Date.objects.get(journal=journal)
                due_days = date_obj.due_days_to_minor_revision
                submission.revision_due_date=timezone.now().date() + timedelta(days=due_days)
            if decision.decision == 'Major Revision':
                journal = submission.journal  
                date_obj = Date.objects.get(journal=journal)
                due_days = date_obj.due_days_to_major_revision
                submission.revision_due_date=timezone.now().date() + timedelta(days=due_days)
                
            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                        to_email=submission.author.user.email,
                        subject='Manuscript Decision',
                        template_name='email_templates/decision_revision.html',
                        user=submission.author.user,
                        context={'submission': submission }
                    )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=submission.author,
                    template_name='email_templates/decision_revision.html',
                    context={'submission': submission }
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=submission.author,
                    template_name='email_templates/decision_revision.html',
                    context={'submission': submission }
                )
        else:
            if decision.decision == 'Accepted':
                journal = submission.journal  
                date_obj = Date.objects.get(journal=journal)
                due_days = date_obj.due_days_to_payment
                submission.payment_due_date=timezone.now().date() + timedelta(days=due_days)
                if Modes.objects.filter(name="Email",is_active=True):
                    send_email(
                        to_email=submission.author.user.email,
                        subject='Manuscript Decision',
                        template_name='email_templates/decision_accepted.html',
                        user=submission.author.user,
                        context={'submission': submission }
                    )
                if Modes.objects.filter(name="Whatsapp",is_active=True):
                    send_whatsapp(
                        user=submission.author,
                        template_name='email_templates/decision_accepted.html',
                        context={'submission': submission }
                    )
                if Modes.objects.filter(name="SMS",is_active=True):
                    send_sms(
                        user=submission.author,
                        template_name='email_templates/decision_accepted.html',
                        context={'submission': submission }
                    )
            else:
                if Modes.objects.filter(name="Email",is_active=True):
                    send_email(
                        to_email=submission.author.user.email,
                        subject='Manuscript Decision',
                        template_name='email_templates/decision_rejected.html',
                        user=submission.author.user,
                        context={'submission': submission }
                    )
                if Modes.objects.filter(name="Whatsapp",is_active=True):
                    send_whatsapp(
                        user=submission.author,
                        template_name='email_templates/decision_rejected.html',
                        context={'submission': submission }
                    )
                if Modes.objects.filter(name="SMS",is_active=True):
                    send_sms(
                        user=submission.author,
                        template_name='email_templates/decision_rejected.html',
                        context={'submission': submission }
                    )
            submission.article_status=Article_Status.objects.get(article_status=decision.decision)
        
        submission.decissioned_on = timezone.now()  
        submission.save()

        return redirect('editor_in_chief')  # Redirect to the editor in chief page after submission

    return render(request, 'editor_in_chief.html')

@login_required
@user_passes_test(is_eic)
def decisioned_manuscripts(request):
    # Get all unique article statuses
    all_statuses = Submission.objects.exclude(article_status__article_status='Draft').values_list('article_status__article_status', flat=True).distinct()

    # Get the search term from the request
    search_term = request.GET.get('search', '')
    
    # Filter submissions based on selected status and search term
    selected_status = request.GET.get('status', None)
    
        # Get the logged-in user's journal
    user = request.user
    try:
        editor = Editor.objects.get(user=user)
        journals = Journal_Editor_Assignment.objects.filter(editor=editor).values_list('journal', flat=True)
    except Editor.DoesNotExist:
        journals = []

    # Filter submissions by article status and the logged-in user's journal
    submissions = Submission.objects.filter(article_status__article_status__in=[
            'AE Assigned', 'Awaiting for Reviewers', 'Published', 'Rejected',
            'Awaiting for Proof Read', 'Awaiting for Revision', 'In Admin Processing',
            'Awaiting for Correction after Acceptance', 'Payment Done',
            'Awaiting for EIC Decision', 'Corrected', 'Awaiting Changes',
            'Ready to Publish', 'Type Set Done', 'Under Review', 'Submitted',
            'Proof Read Done', 'Awaiting AE Recommendation', 'Accepted'
        ],journal__in=journals,is_decissioned=True).order_by('-id')
    
    
    if selected_status:
        submissions = submissions.filter(article_status__article_status=selected_status,is_decissioned=True).order_by('-id')

    if search_term:
        submissions = submissions.filter(title__icontains=search_term,is_decissioned=True).order_by('-id')

    paginator = Paginator(submissions, 20)  # Show 20 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)

    return render(request, 'decisioned_manuscripts.html', {
        'submissions': submissions,
        'all_statuses': all_statuses,
        'selected_status': selected_status,
        'search_term': search_term,
        'feedback_types': FeedbackType.objects.all(),
        'all_questions': Question.objects.all(),
    })


# -----------------------------------------------------------------------Reviewers-----------------------------------------------------------------------------------------------------------------------------------------

logger = logging.getLogger(__name__)
@login_required
@user_passes_test(is_reviewer)

def accept_invitation(request):
    try:
        if request.method == 'POST':
            invite_id = request.POST.get('invite_id')
            logger.debug(f"Invite ID: {invite_id}")
            invitation = Reviewer_Invitation.objects.get(id=invite_id)
            logger.debug(f"Invitation: {invitation}")
            author = Author.objects.get(user=invitation.user)
            invitation.invite_status = 'A'
            journal = invitation.submission.journal
            date_obj = Date.objects.get(journal=journal)
            due_days = date_obj.due_days_to_review
            logger.debug(f"Due Days: {due_days}")
            invitation.submission.revision_due_date = now() + timedelta(days=due_days)
            invitation.save()
            request_status = Request_Status.objects.get(request_status='Assigned')

            # Calculate the actual due_date
            due_date = now() + timedelta(days=due_days)

            submission_reviewer, created = Submission_Reviewer.objects.get_or_create(
                submission=invitation.submission,
                reviewer=author,
                defaults={
                    'accepted_on': now(),
                    'request_status': request_status,
                    'review_recommendation': '',  # Set a default value or handle it appropriately
                    'review_comments': '',
                    'due_date': due_date  # Use the calculated due_date
                }
            )

            if not created:
                logger.info(f"Submission Reviewer already exists: {submission_reviewer}")

            submission = invitation.submission

            # AE (editor) assigned mail
            try:
                ae_assignment = AE_Assignment.objects.get(submission=submission)
                editor = ae_assignment.user
                if Modes.objects.filter(name="Skip AE",is_active=True):
                    editor = invitation.invite_by.user
            except AE_Assignment.DoesNotExist:
                editor = None
                logger.error("No AE (editor) assigned to this submission.")
                return JsonResponse({'status': 'Error', 'message': 'No AE (editor) assigned to this submission.'}, status=500)
            
            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                    to_email=editor.email,
                    subject='Invitation Accepted',
                    template_name='email_templates/reviewer_accepted.html',
                    user=editor,
                    context={'submission': submission, 'reviewer': author, 'editor': editor}
                )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=editor,
                    template_name='email_templates/reviewer_accepted.html',
                    context={'submission': submission, 'reviewer': author, 'editor': editor}
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=editor,
                    template_name='email_templates/reviewer_accepted.html',
                    context={'submission': submission, 'reviewer': author, 'editor': editor}
                )

            all_accepted = (
                Reviewer_Invitation.objects.filter(submission=submission, invite_status='A').count() ==
                Reviewer_Invitation.objects.filter(submission=submission).count()
            )
            logger.debug(f"All Accepted: {all_accepted}")

            if all_accepted:
                submission.article_status = Article_Status.objects.get(article_status='Under Review')
                submission.save()

            return JsonResponse({'status': 'Accepted'})

    except Exception as e:
        logger.error(f"Error in accept_invitation: {e}")
        return JsonResponse({'status': 'Error', 'message': str(e)}, status=500)


def reject_invitation(request):
    if request.method == 'POST':
        invite_id = request.POST.get('invite_id')
        invitation = Reviewer_Invitation.objects.get(id=invite_id)
        author = Author.objects.get(user=invitation.user)
        invitation.delete()
        submission = invitation.submission
        try:
            ae_assignment = AE_Assignment.objects.get(submission=submission)
            editor = ae_assignment.user
            if Modes.objects.filter(name="Skip AE",is_active=True):
                    editor = invitation.invite_by.user
        except AE_Assignment.DoesNotExist:
            editor = None
            logger.error("No AE (editor) assigned to this submission.")
            return JsonResponse({'status': 'Error', 'message': 'No AE (editor) assigned to this submission.'}, status=500)
        
        if Modes.objects.filter(name="Email",is_active=True):
            send_email(
                to_email=editor.email,
                subject='Invitation Rejected',
                template_name='email_templates/reviewer_rejected.html',
                user=editor,
                context={'submission': submission, 'reviewer': author, 'editor': editor}
            )
        if Modes.objects.filter(name="Whatsapp",is_active=True):
            send_whatsapp(
                user=editor,
                template_name='email_templates/reviewer_rejected.html',
                context={'submission': submission, 'reviewer': author, 'editor': editor}
            )
        if Modes.objects.filter(name="SMS",is_active=True):
            send_sms(
                user=editor,
                template_name='email_templates/reviewer_rejected.html',
                context={'submission': submission, 'reviewer': author, 'editor': editor}
            )
        return JsonResponse({'status': 'Rejected'})
    
@login_required
@user_passes_test(is_reviewer)
def reviewer_invitations(request): 
    author = Author.objects.get(user=request.user)
    reviewer_spl = Reviewer_Specialization.objects.filter(reviewer=request.user)

    # Save specializations if submitted
    if request.method == 'POST' and 'specializations' in request.POST:
        selected_ids = request.POST.getlist('specializations')
        for spec_id in selected_ids:
            specialization = Specialization.objects.get(id=spec_id)
            Reviewer_Specialization.objects.create(reviewer=request.user, specialization=specialization)
        return redirect('reviewer_invitations')  # Reload page after submission

    # Invitations
    invitations = Reviewer_Invitation.objects.filter(user=request.user).exclude(invite_status='RJ').order_by('-id')
    paginator = Paginator(invitations, 5)
    page = request.GET.get('page')

    try:
        invitations = paginator.page(page)
    except PageNotAnInteger:
        invitations = paginator.page(1)
    except EmptyPage:
        invitations = paginator.page(paginator.num_pages)

    all_specializations = Specialization.objects.all()

    return render(request, 'reviewer.html', {
        'invitations': invitations,
        'reviewer_spl': reviewer_spl,
        'all_specializations': all_specializations
    })

# -----------------Manuscripts To Review------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_reviewer)
def Manuscripts_To_Review(request):
    author = Author.objects.get(user=request.user)
    
    # Get Submission IDs for which the reviewer has a request_status of "Assigned"
    assigned_submissions = Submission_Reviewer.objects.filter(
        reviewer=author,
        request_status__request_status='Assigned'
    ).values_list('submission', flat=True)
    
    # Now filter the Submission objects based on these IDs and their article_status
    submissions = Submission.objects.filter(
        id__in=assigned_submissions,
        article_status__article_status='Under Review'
    )
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    
    review_recommendation_choices = Submission_Reviewer.REVIEW_RECOMMENDATION_CHOICES
    
    submission = Submission.objects.all()
    ae_assignments = AE_Assignment.objects.select_related('user').filter(submission__in=submission)
    
    return render(request, 'manuscripts_to_review.html', {'submissions': submissions, 'review_recommendation_choices': review_recommendation_choices,'ae_assignments': ae_assignments, })

from bs4 import BeautifulSoup

def submit_review_comments(request):
    if request.method == 'POST':
        # Get the logged-in user's submissions
        user = Author.objects.get(user=request.user)
        submission_id = request.POST.get('submission_id')

        if submission_id:
            try:
                # Filter the Submission_Reviewer by the logged-in user and the submission_id
                submission_reviewer = Submission_Reviewer.objects.get(submission_id=submission_id, reviewer=user)
            except Submission_Reviewer.DoesNotExist:
                return HttpResponseNotFound("Submission reviewer not found.")
                
            submission_reviewer.review_recommendation = request.POST.get('review_recommendation')
            review_comments_html = request.POST.get('review_comments')
            soup = BeautifulSoup(review_comments_html, "html.parser")
            plain_text_comments = soup.get_text()
            submission_reviewer.review_comments = plain_text_comments
            submission_reviewer.completion_on = now()
            submission_reviewer.request_status=Request_Status.objects.get(request_status='Submitted')
            submission_reviewer.save()
            submission = Submission.objects.get(id=submission_id)
            invitation = Reviewer_Invitation.objects.get(submission=submission, user=user.user)
            try:
                ae_assignment = AE_Assignment.objects.get(submission=submission)
                editor = ae_assignment.user
                if Modes.objects.filter(name="Skip AE",is_active=True):
                    editor = invitation.invite_by.user
            except AE_Assignment.DoesNotExist:
                editor = None
                logger.error("No AE (editor) assigned to this submission.")
                return JsonResponse({'status': 'Error', 'message': 'No AE (editor) assigned to this submission.'}, status=500)

            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                    to_email=editor.email,
                    subject='Invitation Completed',
                    template_name='email_templates/reviewer_completed.html',
                    user=editor,
                    context={'submission': submission, 'reviewer': user, 'editor': editor}
                )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=editor,
                    template_name='email_templates/reviewer_completed.html',
                    context={'submission': submission, 'reviewer': user, 'editor': editor}
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=editor,
                    template_name='email_templates/reviewer_completed.html',
                    context={'submission': submission, 'reviewer': user, 'editor': editor}
                )
            # submission = Submission.objects.get(id=submission_id)
            # submission.article_status = Article_Status.objects.get(article_status='Awaiting AE Recommendation')
                # Check if all Submission_Reviewer entries for this submission are complete
            all_reviews_completed = Submission_Reviewer.objects.filter(submission_id=submission_id).exclude(completion_on=None).count()
            total_reviews = Submission_Reviewer.objects.filter(submission_id=submission_id).count()
            submission = Submission.objects.get(id=submission_id)
            if all_reviews_completed == total_reviews:
                # Update the submission status if all reviews are complete
                # submission = Submission.objects.get(id=submission_id)
                submission.article_status = Article_Status.objects.get(article_status='Awaiting AE Recommendation')
                if Modes.objects.filter(name="Skip AE",is_active=True):
                    submission.article_status = Article_Status.objects.get(article_status='Awaiting for EIC Decision')
            submission.decissioned_on = timezone.now()
            submission.decision_by = request.user
            submission.save()
            return redirect('manuscripts_to_review') 
        else:
            return HttpResponseBadRequest("Invalid submission ID.")
    
    return HttpResponseNotAllowed("Method not allowed.")

# ---------------Manuscripts reviewed--------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_reviewer)
def Reviewed_Manuscripts(request):
    author = Author.objects.get(user=request.user)
    r_subission = Submission_Reviewer.objects.filter(reviewer=author,request_status__request_status="Submitted").values_list('submission',flat=True).order_by('-id')
    submissions=Submission.objects.filter(id__in=r_subission)
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    return render(request, 'reviewed_manuscripts.html', {'submissions': submissions})

# --------------------------------------------------------------------Associate Editor-----------------------------------------------------------------------------------------------------------------------------------------


@login_required
@user_passes_test(is_ae)
def associate_editor(request):
    submissions=Submission.objects.filter(ae_assignment__user=request.user,article_status__article_status__in=['AE Assigned','Awaiting for Reviewers'])
    specializations = Specialization.objects.all()
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    context = {
        'submissions': submissions,
        'specializations': specializations,
        
    }
    return render(request, 'associate_editor.html', context)


@login_required
@user_passes_test(is_ae)
def get_reviewers(request):
    submission_id = request.GET.get('submission_id')
    submission = get_object_or_404(Submission, id=submission_id)
    specialization_id = submission.specialization.id  # Assuming submission.specialization is a FK
    print("Specialization ID:", specialization_id)
    reviewers = User.objects.filter(reviewer_specialization__specialization__id=specialization_id).distinct()
    print("Reviewers found:", reviewers.count())
    reviewers_data = [{'id': r.id, 'name': r.get_full_name()} for r in reviewers]
    return JsonResponse({'reviewers': reviewers_data})



def send_invitation(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            submission_id = data.get('submission_id')
            submission = Submission.objects.get(id=submission_id)
            reviewer_ids = data.get('reviewers')
            journal = submission.journal  
            date_obj = Date.objects.get(journal=journal)
            due_days = date_obj.due_days_to_accept_invitation
            
            if not submission_id or not reviewer_ids:
                return JsonResponse({'success': False, 'error': 'Missing data'}, status=400)

            reviewers = User.objects.filter(id__in=reviewer_ids)
            for reviewer in reviewers:
                # Check if invitation already exists
                if Reviewer_Invitation.objects.filter(user=reviewer, submission=submission).exists():
                    logger.info(f'Invitation already exists for {reviewer.email} for submission {submission.id}')
                    continue

                # Create Reviewer_Invitation
                Reviewer_Invitation.objects.get_or_create(
                    user=reviewer,
                    submission=submission,
                    invite_status='R',
                    expiring_date=timezone.now() + timedelta(days=due_days),
                    invite_by = request.user
                )
                date=timezone.now() + timedelta(days=due_days)
                # Send email
                if Modes.objects.filter(name="Email",is_active=True):
                    send_email(
                        to_email=reviewer.email,
                        subject='Invitation to Review Manuscript',
                        template_name='email_templates/invitation_email.html',
                        user=reviewer,
                        context={'submission': submission, 'reviewer': reviewer,'date':date }
                    )
                if Modes.objects.filter(name="Whatsapp",is_active=True):
                    send_whatsapp(
                        user=reviewer,
                        template_name='email_templates/invitation_email.html',
                        context={'submission': submission, 'reviewer': reviewer,'date':date }
                    )
                if Modes.objects.filter(name="SMS",is_active=True):
                    send_sms(
                        user=reviewer,
                        template_name='email_templates/invitation_email.html',
                        context={'submission': submission, 'reviewer': reviewer,'date':date }
                    )
                
            submission.article_status = Article_Status.objects.get(article_status='Awaiting for Reviewers')
            submission.save()
            
            return JsonResponse({'success': True})
        except Exception as e:
            logger.error(f'Error in send_invitation: {str(e)}')
            return JsonResponse({'success': False, 'error': str(e)}, status=500)
    return JsonResponse({'success': False, 'error': 'Invalid request'}, status=400)


def cancel_invitation(request):
    if request.method == 'POST':
        invite_id = request.POST.get('invite_id')
        Reviewer_Invitation.objects.filter(id=invite_id).delete()
        return JsonResponse({'success': True})

# -----------Manuscripts with Review Reports--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_ae)
def Manuscripts_Review_Report(request):
    submissions = Submission.objects.filter(ae_assignment__user=request.user, article_status__article_status='Awaiting AE Recommendation')
    decisions = Decision.objects.all()
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    admin_group = Group.objects.get(name='Admin Office')

# Get the first user in the 'Admin Office' group (assuming there's only one admin)
    admin_user = admin_group.user_set.first()
    return render(request, 'manuscripts_review_report.html', {
        'submissions': submissions,
        'decisions': decisions,
        'admin_email': admin_user.email if admin_user else '',
    })


def get_reviewers_comments(request):
    if request.method == 'GET':
        comments = Submission_Reviewer.objects.filter(submission_id=request.GET.get('submission_id')).values_list('review_comments', flat=True)
        return JsonResponse({'comments': '\n'.join(comments)})


from django.utils.html import strip_tags
def submit_recommendation(request):
    if request.method == 'POST':
        ae_assignment = AE_Assignment.objects.filter(submission_id=request.POST.get('submission_id')).first()
        if ae_assignment:
            # Strip HTML tags from the comments
            raw_comments = request.POST.get('comments')
            plain_text_comments = strip_tags(raw_comments).replace('\n', '\n\n')
            
            ae_assignment.ae_comments = plain_text_comments
            ae_assignment.ae_recommendation = request.POST.get('recommendation')
            
            submission = Submission.objects.get(id=request.POST.get('submission_id'))
            submission.article_status = Article_Status.objects.get(article_status='Awaiting for EIC Decision')
            submission.save()
            ae_assignment.save()
        
         # Retrieve the Chief Editor (EIC) for the journal
            eic_group = Group.objects.get(name='EIC')
            chief_editor_assignment = Journal_Editor_Assignment.objects.filter(
                journal=submission.journal,
                editor__user__groups=eic_group
            ).first()
            
            if chief_editor_assignment:
                chief_editor = chief_editor_assignment.editor.user
                
                # Send email to Chief Editor
                if Modes.objects.filter(name="Email",is_active=True):
                    send_email(
                        to_email=chief_editor.email,
                        subject='AE Recommendations',
                        template_name='email_templates/ae_completed.html',
                        user=chief_editor,
                        context={'submission': submission, 'chief_editor': chief_editor}
                    )
                if Modes.objects.filter(name="Whatsapp",is_active=True):
                    send_whatsapp(
                        user=chief_editor,
                        template_name='email_templates/ae_completed.html',
                        context={'submission': submission, 'chief_editor': chief_editor}
                    )
                if Modes.objects.filter(name="SMS",is_active=True):
                    send_sms(
                        user=chief_editor,
                        template_name='email_templates/ae_completed.html',
                        context={'submission': submission, 'chief_editor': chief_editor}
                    )
            else:
                # Handle the case where no Chief Editor is found
                return JsonResponse({'status': 'error', 'message': 'Chief Editor not found for this journal.'}, status=404)
            
        return JsonResponse({'status': 'success'})
    return JsonResponse({'status': 'error', 'message': 'Invalid request method.'}, status=400)

    
def get_reviewer_details(request):
    submission_id = request.GET.get('submission_id')
    reviewer_id = request.GET.get('reviewer_id')
    
    try:
        submission_reviewer = Submission_Reviewer.objects.get(submission_id=submission_id, reviewer_id=reviewer_id)
        data = {
            'review_recommendation': submission_reviewer.get_review_recommendation_display(),
            'review_comments': submission_reviewer.review_comments
        }
        return JsonResponse(data)
    except Submission_Reviewer.DoesNotExist:
        return JsonResponse({'error': 'Reviewer not found'}, status=404)
# -----------Manuscripts under review-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_ae)

def manuscripts_under_review(request):
    logged_in_user = request.user
    ae_assignments = AE_Assignment.objects.filter(user=logged_in_user).values_list('submission',flat=True)
    submissions=Submission.objects.filter(id__in=ae_assignments,article_status__article_status='Under Review')
    paginator = Paginator(submissions, 20)  # Show 10 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)
    return render(request,'manuscripts_under_review.html',{'submissions': submissions} )
# --------------manuscrtipts in EIC  ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_ae)

def manuscripts_eic(request):
    logged_in_user = request.user
    ae_assignments = AE_Assignment.objects.filter(user=logged_in_user).values_list('submission',flat=True).order_by('-id')
    submissions=Submission.objects.filter(id__in=ae_assignments)
    # Get all unique article statuses
    all_statuses = Submission.objects.exclude(article_status__article_status='Draft').values_list('article_status__article_status', flat=True).distinct()

    # Get the search term from the request
    search_term = request.GET.get('search', '')
    
    # Filter submissions based on selected status and search term
    selected_status = request.GET.get('status', None)
    
    
    if selected_status:
        submissions = submissions.filter(article_status__article_status=selected_status).order_by('-id')

    if search_term:
        submissions = submissions.filter(title__icontains=search_term).order_by('-id')

    paginator = Paginator(submissions, 20)  # Show 20 submissions per page
    page_number = request.GET.get('page')
    submissions = paginator.get_page(page_number)

    return render(request, 'manuscripts_eic.html', {
        'submissions': submissions,
        'all_statuses': all_statuses,
        'selected_status': selected_status,
        'search_term': search_term,
    })

# -----------Contact for mail-------------------------------------------------------------------------------------------------------
def contact(request):
    to_email = request.GET.get('email', '')  
    if request.method == 'POST':
        form = ContactForm(request.POST)
        if form.is_valid():
            
            subject = form.cleaned_data['subject']
            message = form.cleaned_data['message']
            user =  User.objects.get(email=to_email)
            context = {'message': message}
            
            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                    to_email=to_email,
                    subject=subject,
                    template_name='email_templates/message_form.html',
                    user=user,
                    context=context,
                )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=to_email,
                    template_name='email_templates/message_form.html',
                    context=context,
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=to_email,
                    template_name='email_templates/message_form.html',
                    context=context,
                )
            return redirect(reverse('success_page'))
    else:
        form = ContactForm(to_email=to_email)  
    return render(request, 'contact_form.html', {'form': form})

# -----------success for mail-------------------------------------------------------------------------------------------------------

def success_page(request):
    return render(request, 'success_page.html')


import pandas as pd
from django.http import HttpResponse
from .models import Submission  # Adjust the import based on your project structure

def export_submissions_to_excel(request):
 
    submissions = Submission.objects.filter(article_status__article_status="Accepted")

    data = []
    for submission in submissions:
        submission_file_url = request.build_absolute_uri(submission.final_file.url) if submission.final_file else 'Not Available'
        
        data.append({
            "Manuscript ID": submission.manuscript_id,
            "Title": submission.title,
            "Submission File": submission_file_url,  # Add the submission file link
            "Author": submission.author.get_full_name(),
            "Copyright Form": 'Available' if submission.copyright_file else 'Not Available',
        })

    df = pd.DataFrame(data)
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=submissions.xlsx'
    df.to_excel(response, index=False)

    return response

# dashboard
from django.shortcuts import render
from django.db.models import Count, Max
from .models import Submission  # Adjust import if needed

def dashboard_view(request):
    total_submissions = Submission.objects.count()
    accepted_papers = Submission.objects.filter(decision__decision__iexact="Accepted").count()
    decided_papers = Submission.objects.filter(is_decissioned=True).count()

    # For papers in different stages
    stages = (
        Submission.objects
        .filter(article_status__article_status__in=['Accepted', 'Rejected', 'Submitted', 'Published'])
        .values('article_status__article_status')
        .annotate(count=Count('id'))
    )
    needed = Submission.objects.filter(article_status__article_status__in=['Accepted', 'Rejected', 'Submitted', 'Published']).values('article_status__article_status')
    # Format the data for the chart
    stage_labels = [stage['article_status__article_status'] for stage in stages]
    stage_counts = [stage['count'] for stage in stages]

    # Create a dictionary for the button display
    stage_dict = {stage['article_status__article_status']: stage['count'] for stage in stages}
    country_data = (
        Submission.objects
        .exclude(author__country__isnull=True)
        .values('author__country__country')  # Assuming Country model has a 'name' field
        .annotate(count=Count('id'))
        .order_by('-count')[:10]  # Top 10 countries
    )
    location_data = (
        needed
        .exclude(author__city__isnull=True)
        .exclude(author__city__exact='')
        .values(
            'author__country__country',
            'author__country__code',
            'author__state',
            'author__city'
        )
        .annotate(
            paper_count=Count('id'),
            sample_title=Max('title')
        )
        .order_by('author__country__country', 'author__state', 'author__city')
    )
    
    # Normalize city names and group
    normalized_locations = []
    city_mapping = {
        'kovai': 'Coimbatore',
        'trichy': 'Tiruchirappalli'
    }
    
    for item in location_data:
        city = city_mapping.get(item['author__city'].lower(), item['author__city'])
        normalized_locations.append({
            'country': item['author__country__country'],
            'country_code': item['author__country__code'],
            'state': item['author__state'],
            'city': city,
            'paper_count': item['paper_count'],
            # 'sample_title': item['sample_title']
        })
    # Format for visualization
    countries = [item['author__country__country'] for item in country_data]
    country_counts = [item['count'] for item in country_data]
    # Add this section for reviewer map data
    reviewer_locations = (
        Author.objects
        .filter(is_reviewer=True)
        .exclude(country__isnull=True)
        .values(
            'country__country',
            'country__code',
            'state',
            'city',
            'user__first_name',
            'user__last_name'
        )
        .annotate(count=Count('id'))
    )

    # Normalize reviewer city names using your existing mapping
    normalized_reviewers = []
    city_mapping = {
        'kovai': 'Coimbatore',
        'trichy': 'Tiruchirappalli'
    }
    
    for reviewer in reviewer_locations:
        city = city_mapping.get(reviewer['city'].lower(), reviewer['city'])
        normalized_reviewers.append({
            'country': reviewer['country__country'],
            'country_code': reviewer['country__code'],
            'state': reviewer['state'],
            'city': city,
            'count': reviewer['count'],
            'name': f"{reviewer['user__first_name']} {reviewer['user__last_name']}"
        })

    # Get the earliest and latest submission years for the filter
    # Get the earliest and latest submission years for the filter
    # if Submission.objects.exists() and Submission.objects.exclude(submitted_on__isnull=True).exists():
    #     earliest_submission = Submission.objects.exclude(submitted_on__isnull=True).earliest('submitted_on')
    #     earliest_year = earliest_submission.submitted_on.year
    # else:
    #     earliest_year = datetime.now().year
    # current_year = datetime.now().year
    # years_range = range(earliest_year, current_year + 1)
    
    # # Get selected year from request (default to current year)
    # selected_year = int(request.GET.get('year', current_year))
    
    # # Monthly submissions data (all submissions)
    # monthly_submissions = (
    #     Submission.objects
    #     .filter(submitted_on__year=selected_year)
    #     .annotate(month=ExtractMonth('submitted_on'))
    #     .values('month')
    #     .annotate(count=Count('id'))
    #     .order_by('month')
    # )
    
    # # Monthly accepted papers data
    # monthly_accepted = (
    #     Submission.objects
    #     .filter(
    #         Q(article_status__article_status='Accepted') & 
    #         Q(submitted_on__year=selected_year))
    #     .annotate(month=ExtractMonth('submitted_on'))
    #     .values('month')
    #     .annotate(count=Count('id'))
    #     .order_by('month')
    # )
    
    # # Prepare data for chart - initialize all months with 0
    # month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
    #               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # # Initialize data arrays with zeros
    # submissions_data = [0] * 12
    # accepted_data = [0] * 12
    
    # # Fill in actual data
    # for item in monthly_submissions:
    #     submissions_data[item['month'] - 1] = item['count']
    
    # for item in monthly_accepted:
    #     accepted_data[item['month'] - 1] = item['count']
        
    earliest_year = Submission.objects.earliest('submitted_on').submitted_on.year
    current_year = datetime.now().year
    years_range = range(earliest_year, current_year + 1)

    # Get selected year from request (default to current year)
    selected_year = int(request.GET.get('year', current_year))

  # Monthly submissions data (all submissions)
    monthly_submissions = (
        Submission.objects
        .filter(submitted_on__year=selected_year)
        .annotate(month=ExtractMonth('submitted_on'))
        .values('month')
        .annotate(count=Count('id'))
        .order_by('month')
    )

    # Monthly accepted papers data
    monthly_accepted = (
        Submission.objects.filter(
        Q(article_status__article_status='Accepted') & 
        Q(submitted_on__year=selected_year))
        .annotate(month=ExtractMonth('submitted_on'))
        .values('month')
        .annotate(count=Count('id'))
        .order_by('month')
    )

    # Prepare data for chart - initialize all months with 0
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

    # Initialize data arrays with zeros
    submissions_data = [0] * 12
    accepted_data = [0] * 12

    # Fill in actual data
    for item in monthly_submissions:
        submissions_data[item['month'] - 1] = item['count']

    for item in monthly_accepted:
        accepted_data[item['month'] - 1] = item['count']
    # Get reviewer specialization data
    reviewer_specializations = (
        Reviewer_Specialization.objects
        .values('specialization__specialization')
        .annotate(count=Count('reviewer', distinct=True))
        .order_by('-count')
    )
    
    # Prepare data for the specialization bar chart
    specialization_labels = [spec['specialization__specialization'] for spec in reviewer_specializations]
    specialization_counts = [spec['count'] for spec in reviewer_specializations]
        # Get published articles by specialization data
    published_by_specialization = (
        Published_article.objects
        .exclude(area__isnull=True)
        .values('area__specialization')
        .annotate(count=Count('id'))
        .order_by('-count')
    )
    
    # Prepare data for the published articles specialization bar chart
    pub_spec_labels = [spec['area__specialization'] for spec in published_by_specialization]
    pub_spec_counts = [spec['count'] for spec in published_by_specialization]
    
    context = {
        'total_submissions': total_submissions,
        'accepted_papers': accepted_papers,
        'accepted_percent': round((accepted_papers / total_submissions) * 100, 2) if total_submissions else 0,
        'decided_ratio': round((decided_papers / total_submissions) * 100, 2) if total_submissions else 0,
        'stages': stages,
        'stage_labels': stage_labels,
        'stage_counts': stage_counts,
        'stage_dict': stage_dict,  # ✅ Add this to avoid template error
        'countries': countries,
        'country_counts': country_counts,
        'country_data': country_data,
        'locations': json.dumps(normalized_locations),
        'reviewer_locations': json.dumps(normalized_reviewers),
        'years_range': years_range,
        'selected_year': selected_year,
        'monthly_data': {
            'labels': month_names,
            'submissions': submissions_data,
            'accepted': accepted_data,
        },
        'specialization_data': {
            'labels': specialization_labels,
            'counts': specialization_counts,
        },
        'published_spec_data': {
            'labels': pub_spec_labels,
            'counts': pub_spec_counts,
        },
    }

    return render(request, 'dashboard.html', context)