from django.shortcuts import render, redirect ,get_object_or_404
from django.contrib.auth import login ,logout
from django.contrib.auth.models import User,Group
from django.contrib.auth.decorators import login_required , user_passes_test
from django.contrib.auth import authenticate, login
from django.db import IntegrityError  
from account.sms import send_sms
from account.whatsapp import send_whatsapp
from oss.services import send_email
from .forms import *
from oss.forms import *
from oss.auth import *
from django.http import HttpResponse, JsonResponse
from django.views.decorators.http import require_GET, require_POST
import logging
from django.core.mail import send_mail
from django.contrib.auth.tokens import default_token_generator
from django.urls import reverse
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode
from django.utils.encoding import force_bytes, force_str
from django.conf import settings
from .forms import UserRegistrationForm
from .models import Author, Feedback, FeedbackResponse, Modes, User,Editor
from django.contrib.auth.forms import SetPasswordForm
from oss.models import CoAuthor, Journal, Journal_Editor_Assignment
from django.contrib.auth.forms import SetPasswordForm     #set_new_password
from django.contrib.auth import views as auth_views       #CustomPasswordResetConfirmView
from django.contrib.auth.forms import PasswordResetForm   #CustomPasswordResetConfirmView
from django.contrib.auth.forms import PasswordResetForm   #password_reset_view
from django.contrib.auth.views import PasswordChangeView  #CustomPasswordChangeView
from django.urls import reverse_lazy                      #CustomPasswordChangeView
from django.contrib.auth.mixins import LoginRequiredMixin 
import json
from django.core.paginator import Paginator
from dl.models import Published_article, Issue
from dl .models import *



# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#**************logout****************

def custom_logout(request):
    logout(request)
    return redirect('login')

# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

logger = logging.getLogger(__name__)

def login_view(request):
    if request.user.is_authenticated:
        user = request.user
        if user.is_superuser:
            return redirect('dashboard')  # Redirect to user_management.html
        elif user.groups.filter(name='Admin Office').exists():
            return redirect('dashboard')  # Redirect to admin_office.html
        elif user.groups.filter(name='AE').exists():
            return redirect('associate_editor')  # Redirect to AE dashboard or specific view
        elif user.groups.filter(name='EIC').exists():
            return redirect('editor_in_chief')  # Redirect to EIC dashboard or specific view
        elif user.groups.filter(name='Author').exists():
            return redirect('startnew')  # Redirect to Author dashboard or specific view
        elif user.groups.filter(name='Reviewer').exists():
            return redirect('reviewer_invitations')  # Redirect to Reviewer dashboard or specific view
        else:
            return redirect('change_password')  # Redirect for other users

    if request.method == 'POST':
        form = LoginForm(request.POST)
        if form.is_valid():
            username = form.cleaned_data['username']
            password = form.cleaned_data['password']
            user = authenticate(request, username=username, password=password)
            if user is not None:
                login(request, user)
                if user.is_superuser:
                    return redirect('dashboard')  # Redirect to user_management.html
                elif user.groups.filter(name='Admin Office').exists():
                    return redirect('dashboard')  # Redirect to admin_office.html
                elif user.groups.filter(name='AE').exists():
                    return redirect('associate_editor')  # Redirect to AE dashboard or specific view
                elif user.groups.filter(name='EIC').exists():
                    return redirect('editor_in_chief')  # Redirect to EIC dashboard or specific view
                elif user.groups.filter(name='Author').exists():
                    return redirect('startnew')  # Redirect to Author dashboard or specific view
                elif user.groups.filter(name='Reviewer').exists():
                    return redirect('reviewer_invitations')  # Redirect to Reviewer dashboard or specific view
                else:
                    return redirect('change_password')  # Redirect to a general view for other users
            else:
                form.add_error(None, 'Invalid username or password')
        else:
            form.add_error(None, 'Form is not valid')
    else:
        form = LoginForm()

    return render(request, 'login.html', {'form': form})
 
 
# -------------------------------------------------------------------------------------user_management--------------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_super_admin)
def user_management(request):
    users_list = User.objects.all().order_by('-id')

    paginator = Paginator(users_list,20)  # Show 20 users per page.
    page_number = request.GET.get('page')
    users = paginator.get_page(page_number)
    
    journal_usernames = list(User.objects.filter(groups__name__in=['AE', 'EIC']).values_list('username', flat=True))
    dates = Date.objects.all()
    context = {
        'users': users,
        'journal_usernames': json.dumps(journal_usernames), 
        'journals' : Journal.objects.all() ,# Convert list to JSON string
        'dates': dates,
    }
    return render(request, 'user_management.html', context)


# --------------------Create user------------------------------------------------------------------------------------------------------------------------------------------------------------------------


@login_required
@user_passes_test(is_super_admin)
def create_user(request):
    if request.method == 'POST':
        first_name = request.POST.get('first_name')
        last_name = request.POST.get('last_name')
        email = request.POST.get('email')
        is_active = request.POST.get('is_active') == 'true'

        if not email:
            return JsonResponse({'success': False, 'error': 'Email is required'})

        username = email  # Set the username as the email

        try:
            temp_password = get_random_string(length=8)  # Generate a random temporary password
            user = User.objects.create_user(
                username=username,
                email=email,
                first_name=first_name,
                last_name=last_name,
                password=temp_password,  # Use the generated random password
                is_active=is_active
            )
            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                        to_email=user.email,
                        subject='Your account has been created',
                        template_name='email_templates/account_created.html',
                        user=user,
                        context={'user': user }
                    )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=user,
                    template_name='email_templates/account_created.html',
                    context={'user': user}
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=user,
                    template_name='email_templates/account_created.html',
                    context={'user': user}
                )
            
            return JsonResponse({'success': True})
        except Exception as e:
            return JsonResponse({'success': False, 'error': str(e)})
    return JsonResponse({'success': False, 'error': 'Invalid request method'})

# --------------------Assign Roles------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# @require_GET
# def fetch_groups(request):
#     username = request.GET.get('username')
#     user = User.objects.get(username=username)
#     groups = Group.objects.all()
#     user_groups = user.groups.values_list('id', flat=True)
#     groups_data = [{'id': group.id, 'name': group.name} for group in groups]
#     return JsonResponse({'groups': groups_data, 'user_groups': list(user_groups)})



# @require_POST
# def update_user_groups(request):
#     username = request.POST.get('username')
#     groups = request.POST.getlist('groups[]')
#     user = get_object_or_404(User, username=username)
#     user.groups.clear()
#     ae_group = Group.objects.get(name='AE')
#     eic_group = Group.objects.get(name='EIC')
#     admin_group = Group.objects.get(name='Admin Office')
#     verification_link = f"{settings.SITE_URL}{reverse('login')}"
#     for group_id in groups:
#         group = get_object_or_404(Group, id=group_id)
#         user.groups.add(group)

#     # Collect all assigned groups
#     assigned_groups = []

#     # Assign new groups and collect the group names
#     for group_id in groups:
#         group = get_object_or_404(Group, id=group_id)
#         user.groups.add(group)
#         assigned_groups.append(group.name)

#     # Send a single email with all roles assigned
#     if assigned_groups:
#         send_email(
#             to_email=user.email,
#             subject='Roles assigned to you',
#             template_name='email_templates/admin_role.html',
#             user=user,
#             context={'user': user, 'assigned_groups': assigned_groups,'verification_link':verification_link}
#         )
#     # Check if user has AE or EIC roles and update Editor model
#     if ae_group in user.groups.all() or eic_group in user.groups.all():
#         editor, created = Editor.objects.get_or_create(user=user)
#         editor.is_active = True
#         editor.save()
#     else:
#         # If user doesn't have AE or EIC roles, remove from Editor model
#         Editor.objects.filter(user=user).delete()

#     return JsonResponse({'status': 'success'})
from django.contrib.auth.models import User, Group
from django.http import JsonResponse
from django.views.decorators.http import require_GET, require_POST
from django.shortcuts import get_object_or_404
from oss.models import Specialization, Reviewer_Specialization
# from your_app.email_utils import send_email  # Adjust import based on your setup
from django.conf import settings
from django.urls import reverse

@require_GET
def fetch_groups(request):
    username = request.GET.get('username')
    user = User.objects.get(username=username)
    groups = Group.objects.all()
    user_groups = user.groups.values_list('id', flat=True)
    groups_data = [{'id': group.id, 'name': group.name} for group in groups]
    
    # Fetch specializations
    specializations = Specialization.objects.all()
    specializations_data = [{'id': spec.id, 'specialization': spec.specialization} for spec in specializations]
    print("Specializations:", specializations_data)
    # Get Reviewer group ID
    reviewer_group = Group.objects.get(name='Reviewer')
    
    return JsonResponse({
        'groups': groups_data,
        'user_groups': list(user_groups),
        'specializations': specializations_data,
        'reviewer_group_id': reviewer_group.id
    })

@require_GET
def fetch_reviewer_specializations(request):
    username = request.GET.get('username')
    user = User.objects.get(username=username)
    reviewer_specs = Reviewer_Specialization.objects.filter(reviewer=user).values_list('specialization__id', flat=True)
    return JsonResponse({'specialization_ids': list(reviewer_specs)})

@require_POST
def update_user_groups(request):
    username = request.POST.get('username')
    groups = request.POST.getlist('groups[]')  # Changed to groups[] to match jQuery serialization
    specializations = request.POST.getlist('specializations[]')  # Get specializations
    user = get_object_or_404(User, username=username)
    
    # Clear existing groups
    user.groups.clear()
    
    ae_group = Group.objects.get(name='AE')
    eic_group = Group.objects.get(name='EIC')
    admin_group = Group.objects.get(name='Admin Office')
    reviewer_group = Group.objects.get(name='Reviewer')
    verification_link = f"{settings.SITE_URL}{reverse('login')}"
    
    # Assign new groups
    assigned_groups = []
    for group_id in groups:
        group = get_object_or_404(Group, id=group_id)
        user.groups.add(group)
        author_id = Author.objects.get(user=user) if Author.objects.filter(user=user).exists() else None
        if author_id is None:
            if group.name == 'Reviewer':
                Author.objects.create(
                user=user,
                title=Title.objects.all().first(),  
                institution='update',
                address='update',
                city='update',
                state='update',
                country=Country.objects.get(country='India'),  
                mobile_no='update',
                is_reviewer=True,
                )
        assigned_groups.append(group.name)
        
        if group.name == admin_group.name:
            old_admins = User.objects.filter(groups=admin_group).exclude(id=user.id)
            for old_admin in old_admins:
                # Remove admin group from old admin user
                if old_admin: 
                    # Send email notifying old user their role was removed
                    if Modes.objects.filter(name="Email",is_active=True):
                        send_email(
                            to_email=old_admin.email,
                            subject='Admin Office Role Removed',
                            template_name='email_templates/admin_role_removed.html',
                            user=old_admin,
                            context={'user': old_admin}
                        )
                    if Modes.objects.filter(name="Whatsapp",is_active=True):
                        send_whatsapp(
                            user=old_admin,
                            template_name='email_templates/admin_role_removed.html',
                            context={'user': old_admin}                            
                        )
                    if Modes.objects.filter(name="SMS",is_active=True):
                        send_sms(
                            user=old_admin,
                            template_name='email_templates/admin_role_removed.html',
                            context={'user': old_admin}                            
                        )
                    old_admin.groups.remove(admin_group)
        
    
    # Send email with all assigned roles
    if assigned_groups:
        if Modes.objects.filter(name="Email",is_active=True):
            send_email(
                to_email=user.email,
                subject='Roles assigned to you',
                template_name='email_templates/admin_role.html',
                user=user,
                context={'user': user, 'assigned_groups': assigned_groups, 'verification_link': verification_link}
            )
        if Modes.objects.filter(name="Whatsapp",is_active=True):
            send_whatsapp(
                user=user,
                template_name='email_templates/admin_role.html',
                context={'user': user, 'assigned_groups': assigned_groups, 'verification_link': verification_link}
            )
        if Modes.objects.filter(name="SMS",is_active=True):
            send_sms(
                user=user,
                template_name='email_templates/admin_role.html',
                context={'user': user, 'assigned_groups': assigned_groups, 'verification_link': verification_link}
            )
    # Update Editor model for AE/EIC roles
    if ae_group in user.groups.all() or eic_group in user.groups.all():
        editor, created = Editor.objects.get_or_create(user=user)
        editor.is_active = True
        editor.save()
    else:
        Editor.objects.filter(user=user).delete()

    return JsonResponse({'status': 'success'})


@require_POST
def add_specialization(request):
    new_specialization = request.POST.get('newSpecialization')
    if not new_specialization:
        return JsonResponse({'success': False, 'error': 'Specialization name is required.'})

    # Create and save the new specialization
    specialization = Specialization(specialization=new_specialization)
    specialization.save()

    return JsonResponse({'success': True})

# @require_GET
# def fetch_specializations(request):
#     specializations = Specialization.objects.all()
#     specializations_data = [{'id': spec.id, 'specialization': spec.specialization} for spec in specializations]
#     return JsonResponse({'specializations': specializations_data})

# --------------------Reset User Password------------------------------------------------------------------------------------------------------------------------------------------------------------------------

logger = logging.getLogger(__name__)
@login_required
@user_passes_test(is_super_admin)
def reset_user_password(request):
    if request.method == 'POST':
        try:
            username = request.POST['username']
            new_password = request.POST['new_password']
            email = request.POST['email']
            
            user = User.objects.get(username=username)
            user.set_password(new_password)
            user.save()
            verification_link = f"{settings.SITE_URL}{reverse('change_password')}"

            context = {
                'username': username,
                'new_password': new_password,
                'verification_link': verification_link,
            }

            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                    to_email=email,
                    subject='Your password has been reset',
                    template_name='email_templates/reset_password_email.html',
                    user=user,
                    context=context
                )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=user,
                    template_name='email_templates/reset_password_email.html',
                    context=context
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=user,
                    template_name='email_templates/reset_password_email.html',
                    context=context
                )
            return JsonResponse({'success': True})
        except User.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User does not exist'}, status=404)
        except Exception as e:
            logger.error(f"Error occurred: {str(e)}")  # Log the error
            return JsonResponse({'success': False, 'error': str(e)}, status=500)
    return JsonResponse({'success': False, 'error': 'Invalid request method'}, status=400)

# --------------------Journal Assign------------------------------------------------------------------------------------------------------------------------------------------------------------------------

def get_journals(request):
    username = request.GET.get('username')
    editor = Editor.objects.filter(user__username=username).first()
    
    journals = Journal.objects.all()
    assigned_journal = Journal_Editor_Assignment.objects.filter(editor=editor).first()
    
    data = {
        'journals': list(journals.all().values('id', 'title')),
        'selected_journal': assigned_journal.journal.id if assigned_journal else None
    }
    
    return JsonResponse(data)
@login_required
@user_passes_test(is_super_admin)
def assign_journal(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        journal_id = request.POST.get('journal_id')

        if not username or not journal_id:
            return JsonResponse({'success': False, 'error': 'Username or Journal ID is missing.'})

        try:
            user = get_object_or_404(User, username=username)
            journal = get_object_or_404(Journal, id=journal_id)
            editor, created = Editor.objects.get_or_create(user=user)

            # Check if user is Editor-in-Chief
            if editor.user.groups.filter(name='EIC').exists():
                # 1. Remove any existing journal from this EIC
                Journal_Editor_Assignment.objects.filter(editor=editor).delete()

                # 2. Remove any existing EIC assignment from this journal
                existing_assignments = Journal_Editor_Assignment.objects.filter(journal=journal)
                for assignment in existing_assignments:
                    if assignment.editor.user.groups.filter(name='EIC').exists():
                        assignment.delete()

            else:
                # For non-EICs: just remove existing assignments to this journal by the same user
                Journal_Editor_Assignment.objects.filter(editor=editor, journal=journal).delete()

            # Assign journal to the editor
            Journal_Editor_Assignment.objects.create(journal=journal, editor=editor)

            verification_link = f"{settings.SITE_URL}{reverse('login')}"
            
            context = {
                'username': username,
                'journal_name': journal.title,
                'verification_link':verification_link,
            }

            # Send email
            if Modes.objects.filter(name="Email",is_active=True):
                send_email(
                    to_email=user.email,
                    subject='Regarding Journal Assignment',
                    template_name='email_templates/journal_assignment_email.html',
                    user=user,
                    context=context
                )
            if Modes.objects.filter(name="Whatsapp",is_active=True):
                send_whatsapp(
                    user=user,
                    template_name='email_templates/journal_assignment_email.html',
                    context=context
                )
            if Modes.objects.filter(name="SMS",is_active=True):
                send_sms(
                    user=user,
                    template_name='email_templates/journal_assignment_email.html',
                    context=context
                )
            return JsonResponse({'success': True})
        except User.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User does not exist.'})
        except Journal.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'Journal does not exist.'})
        except Exception as e:
            return JsonResponse({'success': False, 'error': str(e)})
    return JsonResponse({'success': False, 'error': 'Invalid request method.'})

# -----------------------------------------------------------------------------------Publication Management---------------------------------------------------------------------------------------------------------


@login_required
@user_passes_test(is_super_admin)
def publication_management(request):
    context = {
        'journals' : Journal.objects.all() ,# Convert list to JSON string
    }
    return render(request, 'publication_management.html', context)

def search_users(request):
    query = request.GET.get('query', '')
    users = User.objects.filter(first_name__icontains=query) | User.objects.filter(username__icontains=query)
    users_data = list(users.values('first_name', 'last_name', 'username', 'email', 'is_superuser'))

    journal_usernames = list(User.objects.filter(groups__name__in=['AE', 'EIC']).values_list('username', flat=True))
    return JsonResponse({'users': users_data, 'journal_usernames': journal_usernames})


# --------------------------------------------------------------------------------------Date Settings---------------------------------------------------------------------------------------
@login_required
@user_passes_test(is_super_admin)
def date_settings(request):
    if request.method == 'POST':
        journal_id = request.POST.get('journal_id')
        journal = get_object_or_404(Journal, id=journal_id)
        date_instance, created = Date.objects.get_or_create(journal=journal)

        date_instance.due_days_to_accept_invitation = request.POST.get('due_days_to_accept_invitation')
        date_instance.due_days_to_review = request.POST.get('due_days_to_review')
        date_instance.due_days_to_minor_revision = request.POST.get('due_days_to_minor_revision')
        date_instance.due_days_to_major_revision = request.POST.get('due_days_to_major_revision')
        date_instance.due_days_to_payment = request.POST.get('due_days_to_payment')
        date_instance.due_days_to_corrections = request.POST.get('due_days_to_corrections')
        date_instance.due_days_to_typeset_approval = request.POST.get('due_days_to_typeset_approval')
        date_instance.due_days_to_next_step = request.POST.get('due_days_to_next_step')
        date_instance.save()

        return redirect('date_settings')  # Redirect to the same page after submission

    journals = Journal.objects.all()
    dates = {date.journal_id: date for date in Date.objects.all()}
    return render(request, 'date_settings.html', {'journals': journals, 'dates': dates})

# ------------------------------------------------------------------------Modes------------------------------------------------------------------------------------------------
def modes(request):
    communication_modes = Modes.objects.all().exclude(name='Skip AE')
    flow_change_modes = Modes.objects.filter(name='Skip AE')
    
    if request.method == 'POST':
        # Get the active tab from POST data
        active_tab = request.POST.get('active_tab', 'communication')
        
        # Update  modes
        for mode in communication_modes:
            mode.is_active = (f'mode_{mode.id}' in request.POST)
            mode.save()

        # Update flow change modes
        for mode in flow_change_modes:
            mode.is_active = (f'mode_{mode.id}' in request.POST)
            mode.save()

        # Redirect with the active tab parameter
        return redirect(f'{reverse("modes")}?active_tab={active_tab}')
    
    # GET request handling
    active_tab = request.GET.get('active_tab', 'communication')
    
    return render(request, 'modes.html', {
        'communication_modes': communication_modes,
        'flow_change_modes': flow_change_modes,
        'active_tab': active_tab,
    })
# ------------------------------------------------------------------------Index home page ---------------------------------------------------------------------------------------
from django.db.models import Q

def index(request):
    # if request.user.is_authenticated:
    #     user = request.user
    #     if user.is_superuser:
    #         return redirect('dashboard')  # Redirect to user_management.html
    #     elif user.groups.filter(name='Admin Office').exists():
    #         return redirect('dashboard')  # Redirect to admin_office.html
    #     elif user.groups.filter(name='AE').exists():
    #         return redirect('associate_editor')  # Redirect to AE dashboard or specific view
    #     elif user.groups.filter(name='EIC').exists():
    #         return redirect('editor_in_chief')  # Redirect to EIC dashboard or specific view
    #     elif user.groups.filter(name='Author').exists():
    #         return redirect('startnew')  # Redirect to Author dashboard or specific view
    #     elif user.groups.filter(name='Reviewer').exists():
    #         return redirect('reviewer_invitations')  # Redirect to Reviewer dashboard or specific view
    #     else:
    #         return redirect('change_password')
        
    journals = Journal.objects.all()
    # journal_name = "JCS"  # The name of the journal to filter by
    journal = Journal.objects.get(title="JCS")  # Assuming 'name' is a unique field in Journal model
    
    volumes = Volume.objects.filter(journal=journal)
    latest_issues = Issue.objects.filter(volume__journal=journal).order_by('-id')[:6]  # Fetch the latest 3 issues for the journal
    
    jcs_journal = Journal.objects.get(title='JCS')  # Adjust according to your Journal model

    # Get the latest volume for the JCS journal
    latest_volume = Volume.objects.filter(journal=jcs_journal).order_by('-year', '-volume').first()
    
    if latest_volume:
        # Get the latest issue for the latest volume
        latest_issue = Issue.objects.filter(volume=latest_volume).order_by('-issue').first()

        if latest_issue:
            # Get the latest published articles for the latest issue
            latest_articles = Published_article.objects.filter(issue=latest_issue).order_by('-id')
        else:
            latest_articles = []
    else:
        latest_articles = []
    user = request.user
    has_ae_or_eic = user.groups.filter(name__in=['AE', 'EIC']).exists()
    has_author_or_reviewer = user.groups.filter(name__in=['Author', 'Reviewer']).exists()
     # Handle search functionality
     
    for article in latest_articles:
        accepted_submission = article.accepted_submission
        if accepted_submission and accepted_submission.submission:
            submission = accepted_submission.submission
            co_authors = CoAuthor.objects.filter(submission=submission)
            article.co_authors = co_authors if co_authors.exists() else None
            article_data = {
                'volume': article.issue.volume.volume,
                'year': article.issue.volume.year,
                'issue': article.issue.issue,
                'title': accepted_submission.corrected_title,
                'authors': submission.author.user.get_full_name(),
                'file': accepted_submission.typeset_file
            }
        else:
            article_data = {
                'volume': article.issue.volume.volume,
                'year': article.issue.volume.year,
                'issue': article.issue.issue,
                'title': article.title,
                'authors': article.author,
                'file': article.file,
            }
        article.article_data = article_data
    search_query = request.GET.get('query', '')
    search_results = []
    if search_query:
        search_results = Published_article.objects.filter(
            Q(title__icontains=search_query)
        ).distinct()

    top_downloaded_articles = Published_article.objects.order_by('-download_count')[:3]
    context = {
        'journal': journal,
        'journals': journals,
        'volumes': volumes,
        'latest_issues': latest_issues,
        'latest_issue': latest_issue,
        'latest_articles': latest_articles,
        'has_ae_or_eic': has_ae_or_eic,
        'has_author_or_reviewer': has_author_or_reviewer,
        'search_results': search_results,
        'search_query': search_query,
        'top_downloaded_articles': top_downloaded_articles,
        'latest_volume':latest_volume,
    }
    return render(request, 'index.html', context)

def home(request):
    if request.user.is_authenticated:
        user = request.user
        if user.is_superuser:
            return redirect('dashboard')  # Redirect to user_management.html
        elif user.groups.filter(name='Admin Office').exists():
            return redirect('dashboard')  # Redirect to admin_office.html
        elif user.groups.filter(name='AE').exists():
            return redirect('associate_editor')  # Redirect to AE dashboard or specific view
        elif user.groups.filter(name='EIC').exists():
            return redirect('editor_in_chief')  # Redirect to EIC dashboard or specific view
        elif user.groups.filter(name='Author').exists():
            return redirect('startnew')  # Redirect to Author dashboard or specific view
        elif user.groups.filter(name='Reviewer').exists():
            return redirect('reviewer_invitations')  # Redirect to Reviewer dashboard or specific view
        else:
            return redirect('change_password')
    else:
        return redirect('index')
    return render(request, 'home.html')

def check_user_status(request):
    user = request.user
    return JsonResponse({
        'is_authenticated': user.is_authenticated,
        'is_author': user.groups.filter(name='Author').exists(),
    })

# -------register------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

from django.db import IntegrityError  

def register(request):
    if request.method == 'POST':
        form = UserRegistrationForm(request.POST)
        if form.is_valid():
            # Process the valid form data
            user = form.save(commit=False)
            try:
                user.username = form.cleaned_data.get('email')
                user.is_active = False
                user.save()
                user.refresh_from_db()
            
                phone_code = request.POST.get('phone_code')
                mobile_no = form.cleaned_data.get('mobile_no')
                full_mobile_no = f"{phone_code}{mobile_no}"
            
                author = Author(
                    user=user,
                    title=form.cleaned_data.get('title'),
                    institution=form.cleaned_data.get('institution'),
                    address=form.cleaned_data.get('address'),
                    city=form.cleaned_data.get('city'),
                    state=form.cleaned_data.get('state'),
                    country=form.cleaned_data.get('country'),
                    mobile_no=full_mobile_no,
                    zipcode=form.cleaned_data.get('zipcode'),
                    orcid_id=form.cleaned_data.get('orcid_id'),
                    scopus_id=form.cleaned_data.get('scopus_id'),
                )
                author.save()
                token = generate_token(user)
                uid = encode_uid(user)
                verification_link = f"{settings.SITE_URL}{reverse('set_new_password', kwargs={'uidb64': uid, 'token': token})}"
                if Modes.objects.filter(name="Email",is_active=True):
                    send_email(
                            to_email=user.email,
                            subject='Registration Success',
                            template_name='email_templates/registration_success.html',
                            user=user,
                            context={'user': user ,'verification_link': verification_link,}
                        )
                if Modes.objects.filter(name="Whatsapp",is_active=True):
                    send_whatsapp(
                        user=user,
                        template_name='email_templates/registration_success.html',
                        context={'user': user ,'verification_link': verification_link,}
                    )
                if Modes.objects.filter(name="SMS",is_active=True):
                    send_sms(
                        user=user,
                        template_name='email_templates/registration_success.html',
                        context={'user': user ,'verification_link': verification_link,}
                    )
                return redirect('registration_complete')
            except IntegrityError:
                form.add_error('email', 'This email address is already registered.')
       
            # If form is not valid, render the form with errors
            countries = Country.objects.all()
            return render(request, 'register.html', {'form': form, 'countries': countries})
    else:
        form = UserRegistrationForm()
        countries = Country.objects.all()
        return render(request, 'register.html', {'form': form, 'countries': countries})
# ----Mail Verification------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

def generate_token(user):
    return default_token_generator.make_token(user)

def encode_uid(user):
    return urlsafe_base64_encode(force_bytes(user.pk))

def decode_uid(uid):
    return force_str(urlsafe_base64_decode(uid))

def registration_complete(request):
    return render(request, 'registration_complete.html')

def verify_email(request, uidb64, token):
    try:
        uid = force_str(urlsafe_base64_decode(uidb64))
        user = User.objects.get(pk=uid)
    except (TypeError, ValueError, OverflowError, User.DoesNotExist):
        user = None
        logger.error(f'User does not exist or token is invalid: uidb64={uidb64}, token={token}')
    
    if user is not None and default_token_generator.check_token(user, token):
        return redirect('set_new_password', uidb64=uidb64, token=token)
    else:
        logger.error(f'Token verification failed: uidb64={uidb64}, token={token}')
        return render(request, 'email_verification_failed.html')

    
# ------Set New Password----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def set_new_password(request, uidb64, token):
    try:
        uid = force_str(urlsafe_base64_decode(uidb64))
        user = User.objects.get(pk=uid)
    except (TypeError, ValueError, OverflowError, User.DoesNotExist):
        user = None

    if request.method == 'POST' and user is not None and default_token_generator.check_token(user, token):
        form = SetPasswordForm(user, request.POST)
        if form.is_valid():
            form.save()
            user.is_active = True  # Ensure the user is active
            group = get_object_or_404(Group, name='Author')
            user.groups.add(group)
            user.save()
            logger.info(f'Password successfully set for user: {user.username}')
            return redirect('login')
        else:
            logger.error(f'Form errors: {form.errors}')
    else:
        form = SetPasswordForm(user)
        if user is None:
            logger.error(f'User is None, uidb64: {uidb64}, token: {token}')
        if not default_token_generator.check_token(user, token):
            logger.error(f'Token verification failed, uidb64: {uidb64}, token: {token}')

    context = {
        'form': form,
        'uidb64': uidb64,
        'token': token,
    }

    if uidb64:
        try:
            uid = force_str(urlsafe_base64_decode(uidb64))
            user = get_object_or_404(User, pk=uid)
            context['username'] = user.first_name
        except (TypeError, ValueError, OverflowError, User.DoesNotExist):
            context['username'] = None

    return render(request, 'set_new_password.html', context)



# --------Customer Password Change-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



class CustomPasswordResetConfirmView(auth_views.PasswordResetConfirmView):
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['form'] = self.get_form()
        uidb64 = self.kwargs.get('uidb64')
        if uidb64:
            try:
                uid = force_str(urlsafe_base64_decode(uidb64))
                user = get_object_or_404(User, pk=uid)
                context['username'] = user.first_name
            except (TypeError, ValueError, OverflowError, User.DoesNotExist):
                context['username'] = None
                
        
        return context

from django.utils.crypto import get_random_string
def password_reset_view(request):
    if request.method == 'POST':
        form = PasswordResetForm(request.POST)
        if form.is_valid():
            email = form.cleaned_data['email']
            try:
                user = User.objects.get(email=email)
                if not user.is_active:
                    form.add_error('email', 'After registration, you should have received a registration success email. Please click "Active My Account" link in that email to activate your account.')
                    return render(request, 'password_reset_form.html', {'form': form})
                
            except User.DoesNotExist:
                form.add_error('email', 'Email address not found.')
                return render(request, 'password_reset_form.html', {'form': form})

            form.save(request=request)
            
            return redirect('password_reset_done')
    else:
        form = PasswordResetForm()

    return render(request, 'password_reset_form.html', {'form': form})

class CustomPasswordChangeView(LoginRequiredMixin, PasswordChangeView):
    template_name = 'change_password.html'
    success_url = reverse_lazy('password_change_done')
    
    
# ------Profile Page-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


@login_required

def profile(request):
    user = request.user
    try:
        author = Author.objects.get(user=user)
    except Author.DoesNotExist:
        author = None

    if request.method == 'POST':
        user_form = UserForm(request.POST, instance=user)
        author_form = AuthorForm(request.POST, instance=author)

        if user_form.is_valid() and author_form.is_valid():
            user_form.save()
            author_form.save()
            
            # Add the user to the "Author" group if they are not already in it
                
            return redirect('profile')
        else:
            print("User form errors:", user_form.errors)
            print("Author form errors:", author_form.errors)
    else:
        user_form = UserForm(instance=user)
        author_form = AuthorForm(instance=author)
    specializations = None
    if user.groups.filter(name="Reviewer").exists():
        specializations = Reviewer_Specialization.objects.filter(reviewer=user)

    context = {
        'user_form': user_form,
        'author_form': author_form,
        'author': author,
        'specializations': specializations,
    }
    return render(request, 'profile.html', context)


@login_required
def remove_specialization(request, spec_id):
    if request.method == 'POST':
        try:
            spec = Reviewer_Specialization.objects.get(id=spec_id, reviewer=request.user)
            spec.delete()
            return JsonResponse({'status': 'success'})
        except Reviewer_Specialization.DoesNotExist:
            return JsonResponse({'status': 'error', 'message': 'Not found'}, status=404)
    return JsonResponse({'status': 'error', 'message': 'Invalid request'}, status=400)


@login_required
def available_specializations(request):
    user = request.user
    assigned = Reviewer_Specialization.objects.filter(reviewer=user).values_list('specialization_id', flat=True)
    unassigned = Specialization.objects.exclude(id__in=assigned)

    data = [{'id': spec.id, 'specialization': spec.specialization} for spec in unassigned]
    return JsonResponse(data, safe=False)

@login_required
def add_specializations(request):
    if request.method == 'POST':
        spec_ids = request.POST.getlist('specializations')  # match key here
        print("Received specialization IDs:", spec_ids)

        added_specs = []
        for spec_id in spec_ids:
            try:
                spec = Specialization.objects.get(id=spec_id)
                Reviewer_Specialization.objects.get_or_create(reviewer=request.user, specialization=spec)
                added_specs.append({'id': spec.id, 'name': spec.specialization})
            except Specialization.DoesNotExist:
                continue

        return JsonResponse({'status': 'success', 'added_specializations': added_specs})
    
    return JsonResponse({'status': 'error', 'message': 'Invalid request'}, status=400)

@require_POST
@login_required
def create_specialization(request):
    name = request.POST.get("specialization", "").strip()

    if not name:
        return JsonResponse({"error": "Name is required."}, status=400)

    # Create or get the specialization
    spec, created = Specialization.objects.get_or_create(specialization__iexact=name, defaults={"specialization": name})

    # Link to current user if not already linked
    Reviewer_Specialization.objects.get_or_create(reviewer=request.user, specialization=spec)

    return JsonResponse({
        "id": spec.id,
        "name": spec.specialization,
        "created": created
    })
    
def editor_profile(request):
    editor = get_object_or_404(Editor, user=request.user)

    if request.method == 'POST':
        form = EditorProfileForm(request.POST, instance=editor)
        if form.is_valid():
            form.save()
            return redirect('editor_profile')  # Redirect to the same page after saving
    else:
        form = EditorProfileForm(instance=editor)

    return render(request, 'editor_profile.html', {
        'editor': editor,
        'form': form
    })
# -----------------------------------------------------------------------Current Issues----------------------------------------------------------------------------------------------------------------


def current_issue_view(request, journal_id):
    journal = get_object_or_404(Journal, id=journal_id)
    journals = Journal.objects.all()

    # Get the latest volume for the journal
    latest_volume = Volume.objects.filter(journal=journal).order_by('-year', '-volume').first()

    if latest_volume:
        # Get the latest issue from the latest volume
        latest_issue = Issue.objects.filter(volume=latest_volume).order_by('-issue').first()
        
        if latest_issue:
            # Get the published articles for the latest issue
            latest_issues = Published_article.objects.filter(issue=latest_issue).order_by('-id')
        else:
            latest_issues = Published_article.objects.none()
    else:
        latest_issues = Published_article.objects.none()
        
    for article in latest_issues:
        accepted_submission = article.accepted_submission
        if accepted_submission and accepted_submission.submission:
            submission = accepted_submission.submission
            co_authors = CoAuthor.objects.filter(submission=submission)
            article.co_authors = co_authors if co_authors.exists() else None
            article_data = {
                'volume': article.issue.volume.volume,
                'issue': article.issue.issue,
                'title': accepted_submission.corrected_title,
                'authors': submission.author.user.get_full_name(),
                'file': accepted_submission.typeset_file
            }
        else:
            article_data = {
                'volume': article.issue.volume.volume,
                'issue': article.issue.issue,
                'title': article.title,
                'authors': article.author,
                'file': article.file,
            }
        article.article_data = article_data

    return render(request, 'current_issue.html', {'journal': journal, 'latest_issues': latest_issues,'latest_iss': latest_issue, 'latest_volume': latest_volume, 'journals': journals})





def issue_detail_view(request, journal_id, volume_id, issue_id):
    journals = Journal.objects.all()
    journal = get_object_or_404(Journal, pk=journal_id)
    volume = get_object_or_404(Volume, pk=volume_id, journal=journal)
    issue = get_object_or_404(Issue, pk=issue_id, volume=volume)
    articles = Published_article.objects.filter(issue=issue)
    
    
    for article in articles:
        if article.accepted_submission is None:
            article_data = {
                'volume': article.issue.volume.volume,
                'issue': article.issue.issue,
                'title': article.title,
                'authors': article.author,
                'file' : article.file,
            }
        else:
            article_data = {
                'volume': article.issue.volume.volume,
                'issue': article.issue.issue,
                'title': article.accepted_submission.corrected_title,
                'authors': article.accepted_submission.submission.author.user.get_full_name(),
                'file' : article.accepted_submission.typeset_file,
            }
        article.article_data = article_data
        
    jcs_journal = Journal.objects.get(title='JCS') 
    latest_volume = Volume.objects.filter(journal=jcs_journal).order_by('-year', '-volume').first()
    if latest_volume:
        # Get the latest issue for the latest volume
        latest_issue = Issue.objects.filter(volume=latest_volume).order_by('-issue').first()
        
    return render(request, 'issue_detail.html', {'journal': journal, 'volume': volume, 'issue': issue, 'articles': articles, 'journals': journals, 'latest_issue':latest_issue})








# ----------------------------------------------------------Archives view----------------------------------------------------------------------------------------------------------------

def archive_view(request, journal_id):

    journal = get_object_or_404(Journal, pk=journal_id)
    journals= Journal.objects.all()
    volumes = Volume.objects.filter(journal=journal).order_by('-id')
    
    jcs_journal = Journal.objects.get(title='JCS') 
    latest_volume = Volume.objects.filter(journal=jcs_journal).order_by('-year', '-volume').first()
    if latest_volume:
        # Get the latest issue for the latest volume
        latest_issue = Issue.objects.filter(volume=latest_volume).order_by('-issue').first()
    
    
    return render(request, 'archives.html', {'journal': journal, 'volumes': volumes, 'journals': journals,'latest_issue':latest_issue})

# def issue_detail_view(request, journal_id, volume_id, issue_id):
#     journals = Journal.objects.all()
#     journal = get_object_or_404(Journal, pk=journal_id)
#     volume = get_object_or_404(Volume, pk=volume_id, journal=journal)
#     issue = get_object_or_404(Issue, pk=issue_id, volume=volume)
#     articles = Published_article.objects.filter(issue=issue)
    
#     return render(request, 'issue_detail.html', {'journal': journal, 'volume': volume, 'issue': issue, 'articles': articles, 'journals': journals})




# -----------------------------Journal Management-------------------------------------------------------------------------------------------------------------------------------------
def add_journal(request):
    if request.method == 'POST':
        form = JournalForm(request.POST)
        if form.is_valid():
            journal = form.save()
            Date.objects.create(
                journal=journal,
                due_days_to_accept_invitation=14,
                due_days_to_review=14,
                due_days_to_minor_revision=14,
                due_days_to_major_revision=14,
                due_days_to_payment=14,
                due_days_to_corrections=14,
                due_days_to_typeset_approval=14,
                due_days_to_next_step=14,
            )
            return redirect('add_journal')  # Redirect to a desired URL after saving
    else:
        form = JournalForm()
    
    return render(request, 'add_journal.html', {'form': form})



# --------Base for users--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def base_view(request):
    user = request.user
    ae_or_eic = user.groups.filter(name__in=['AE', 'EIC']).exists()
    # print("-------------------------------------",ae_or_eic)
    author_or_reviewer = user.groups.filter(name__in=['Author', 'Reviewer']).exists()
    
    context = {
        'ae_or_eic': ae_or_eic,
        'author_or_reviewer': author_or_reviewer,
    }
    return render(request, 'base.html',context)

# download_count
def download_article(request, article_id):
    # Retrieve the article
    article = get_object_or_404(Published_article, id=article_id)

    # Increment the download count
    article.download_count += 1
    article.save()

    # Redirect to the file URL
    if article.file:
        return redirect(article.file.url)
    else:
        # Handle case where the file doesn't exist
        return redirect('index') 
    
#  feedback


import os
import re
import json
import subprocess
from django.shortcuts import render
from django.http import FileResponse, JsonResponse
from django.core.files.storage import FileSystemStorage
from .forms import DocumentForm
from .extractors import extract_docx, extract_pdf
from .converter import text_to_latex


def upload_file(request):
    """Handle file upload and LaTeX conversion."""
    if request.method == 'POST':
        form = DocumentForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['document']
            fs = FileSystemStorage()
            filename = fs.save(file.name, file)
            filepath = fs.path(filename)

            # Extract text based on file type
            if file.name.endswith('.docx'):
                text = extract_docx(filepath)
            elif file.name.endswith('.pdf'):
                
                text = extract_pdf(filepath)
                
            else:
                text = file.read().decode('utf-8')

            print("Extracted text:\n", text)  # Debug output
            # Clean and normalize the text
            text = text.strip()
            text = re.sub(r'\n{3,}', '\n\n', text)  # Remove extra blank lines
            
            latex_code = text_to_latex(text)
            os.remove(filepath)
            
            # Debug output
            print("Final LaTeX output:\n", latex_code)
            
            return render(request, 'review.html', {
                'latex_code': latex_code,
                'filename': file.name,
                'extracted_text': text,
            })
    else:
        form = DocumentForm()
    return render(request, 'upload.html', {'form': form})


def compile_pdf(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            latex_code = data.get('latex', '')
            if not latex_code:
                return JsonResponse({'error': 'No LaTeX code provided'}, status=400)

            output_dir = 'latex_output'
            os.makedirs(output_dir, exist_ok=True)
            tex_path = os.path.join(output_dir, 'output.tex')
            with open(tex_path, 'w', encoding='utf-8') as f:
                f.write(latex_code)
            print("Generated LaTeX code:\n", latex_code)  # Add this line for debugging

            try:
                # Use xelatex for better Unicode/font/image support
                result = subprocess.run(
                    ['xelatex', '-interaction=nonstopmode',
                     '-output-directory', output_dir, tex_path],
                    capture_output=True,
                    text=True,
                    encoding='utf-8',  # Ensure UTF-8 decoding for output
                    timeout=60
                )
                pdf_path = os.path.join(output_dir, 'output.pdf')
                if os.path.exists(pdf_path):
                    # Optional: Clean up auxiliary files
                    for ext in ['.aux', '.log', '.out', '.toc']:
                        aux_file = os.path.join(output_dir, 'output' + ext)
                        if os.path.exists(aux_file):
                            try:
                                os.remove(aux_file)
                            except Exception:
                                pass
                    pdf_file = open(pdf_path, 'rb')  # <-- Do NOT use 'with'
                    response = FileResponse(
                        pdf_file,
                        content_type='application/pdf'
                    )
                    response['Content-Disposition'] = 'inline; filename="converted.pdf"'
                    return response
                else:
                    error_msg = result.stderr or "PDF generation failed"
                    return JsonResponse({'error': error_msg}, status=500)
            except subprocess.TimeoutExpired:
                return JsonResponse({'error': 'PDF generation timed out'}, status=500)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)
    return JsonResponse({'error': 'Invalid request'}, status=400)

# views.py
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
from django.forms.models import model_to_dict
from .models import FeedbackType, Question, FeedbackQuestion
from .forms import (
    FeedbackTypeForm,
    QuestionForm,
    FeedbackQuestionForm,
    AddQuestionToTypeForm
)
import json
@login_required
def feedback_types(request):
    types = FeedbackType.objects.all().values('id', 'type')
    return JsonResponse(list(types), safe=False)

def feedback_questions(request, feedback_type_id):
    print(f"Fetching questions for feedback type ID: {feedback_type_id}")
    mapped = FeedbackQuestion.objects.filter(feedback_type_id=feedback_type_id).select_related('question')
    
    # Reformat data to match frontend expectations
    questions = [{
        'id': fq.question.id,           # Frontend expects 'id' (not 'question_id')
        'text': fq.question.question    # Frontend expects 'text' (not 'question_text')
    } for fq in mapped]
    
    # Wrap in 'questions' key and return
    return JsonResponse({'questions': questions})  # Remove safe=False

@login_required
@csrf_exempt
@require_http_methods(['POST'])
def remove_feedback_question(request):
    data = json.loads(request.body)
    form = FeedbackQuestionForm(data)
    if form.is_valid():
        try:
            fq = FeedbackQuestion.objects.get(
                feedback_type_id=form.cleaned_data['feedback_type'],
                question_id=form.cleaned_data['question']
            )
            fq.delete()
            return JsonResponse({'success': True})
        except FeedbackQuestion.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'Mapping not found'}, status=400)
    return JsonResponse({'success': False, 'error': 'Invalid data'}, status=400)

@login_required
@csrf_exempt
@require_http_methods(['POST'])
def add_question_to_type(request):
    data = json.loads(request.body)
    form = AddQuestionToTypeForm(data)

    if not form.is_valid():
        return JsonResponse({'success': False, 'error': 'Invalid data'}, status=400)

    feedback_type_id = form.cleaned_data['feedback_type']
    question_id = form.cleaned_data.get('question')
    question_text = form.cleaned_data.get('question_text')

    if question_text and not question_id:
        # Create new question
        question = Question.objects.create(question=question_text)
        question_id = question.id
    elif question_id:
        # Verify question exists
        if not Question.objects.filter(id=question_id).exists():
            return JsonResponse({'success': False, 'error': 'Question does not exist'}, status=400)
    else:
        return JsonResponse({'success': False, 'error': 'Question or text must be provided'}, status=400)

    # Map question to feedback type if not already mapped
    mapping, created = FeedbackQuestion.objects.get_or_create(
        feedback_type_id=feedback_type_id,
        question_id=question_id
    )
    return JsonResponse({
        'success': True,
        'question': {
            'id': mapping.question.id,
            'question': mapping.question.question
        }
    })
@login_required
@csrf_exempt
@require_http_methods(['POST'])
def create_feedback_type(request):
    data = json.loads(request.body)
    type_text = data.get('type', '').strip()

    if not type_text:
        return JsonResponse({'success': False, 'error': 'Type cannot be empty'}, status=400)

    if FeedbackType.objects.filter(type__iexact=type_text).exists():
        return JsonResponse({'success': False, 'exists': True})

    feedback_type = FeedbackType.objects.create(type=type_text)

    return JsonResponse({
        'success': True,
        'feedback_type': {
            'id': feedback_type.id,
            'type': feedback_type.type
        }
    })

from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from django.http import JsonResponse
from .models import FeedbackType, Question, FeedbackQuestion
import json

@csrf_exempt
@require_POST
@login_required
def update_question_for_feedback_type(request):
    data = json.loads(request.body)
    feedback_type_id = data.get('feedback_type_id')
    question_id = data.get('question_id')
    new_text = data.get('new_text', '').strip()

    if not feedback_type_id or not question_id or not new_text:
        return JsonResponse({'success': False, 'error': 'Missing data'}, status=400)

    try:
        feedback_type = FeedbackType.objects.get(id=feedback_type_id)
        question = Question.objects.get(id=question_id)
    except (FeedbackType.DoesNotExist, Question.DoesNotExist):
        return JsonResponse({'success': False, 'error': 'Invalid feedback type or question'}, status=400)

    # Check if question is mapped to other feedback types
    mapped_types_count = FeedbackQuestion.objects.filter(question=question).count()

    if mapped_types_count > 1:
        # Create new question with new_text and map it to this feedback type
        new_question = Question.objects.create(question=new_text)
        # Map new question
        FeedbackQuestion.objects.create(feedback_type=feedback_type, question=new_question)
        # Remove old mapping for this feedback type
        FeedbackQuestion.objects.filter(feedback_type=feedback_type, question=question).delete()

        return JsonResponse({'success': True, 'question': {'id': new_question.id, 'question': new_question.question}})

    else:
        # Just update the existing question text globally
        question.question = new_text
        question.save()
        return JsonResponse({'success': True, 'question': {'id': question.id, 'question': question.question}})
@require_POST
@login_required
@csrf_exempt
def create_feedback(request):
    try:
        data = json.loads(request.body)
        submission_id = data.get('submission_id')
        feedback_type_id = data.get('feedback_type_id')
        
        if not submission_id or not feedback_type_id:
            return JsonResponse({'success': False, 'error': 'Missing required parameters'}, status=400)
            
        submission = Submission.objects.get(id=submission_id)
        author_user = submission.author.user
        
        # Create the feedback record
        feedback = Feedback.objects.create(
            user=author_user,
            feedback_type_id=feedback_type_id,
            created_by=request.user
        )
        
        site_url = f"{settings.SITE_URL}{reverse('feedback_response_form', args=[feedback.id])}"
        if Modes.objects.filter(name="Email",is_active=True):
            send_email(
                    to_email=submission.author.user.email,
                    subject='Requested feedback from publication',
                    template_name='email_templates/feedback_content.html',
                    user=submission.author.user,
                    context={ 'site_url': site_url, 'user': submission.author.user,}
                    )
        
        if Modes.objects.filter(name="Whatsapp",is_active=True):
            send_whatsapp(
                user=submission.author,
                template_name='email_templates/phone_feedback_content.html',
                context={ 'site_url': site_url, 'user': submission.author.user,},
            )
        if Modes.objects.filter(name="SMS",is_active=True):
            send_sms(
                user=submission.author,
                template_name='email_templates/phone_feedback_content.html',
                context={ 'site_url': site_url, 'user': submission.author.user,},
            )

        return JsonResponse({'success': True})

    except Submission.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'Submission not found'}, status=404)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

# feedback response form view
 # Optional: If you want to require login

def feedback_response_form(request, feedback_id):
    try:
        feedback = Feedback.objects.get(id=feedback_id)
    except Feedback.DoesNotExist:
        return HttpResponse("Invalid feedback link.", status=404)

    if feedback.is_active:
        return render(request, 'thank_you.html', {'user': feedback.user})

    questions = FeedbackQuestion.objects.filter(feedback_type=feedback.feedback_type)
    feedback_options = FeedbackOptions.objects.all().order_by('-value')

    if request.method == "POST":
        for fq in questions:
            selected_option_id = request.POST.get(f'question_{fq.question.id}', '').strip()
            if selected_option_id:
                try:
                    selected_option = FeedbackOptions.objects.get(id=selected_option_id)
                    FeedbackResponse.objects.create(
                        feedback=feedback,
                        question=fq.question,
                        options=selected_option
                    )
                except FeedbackOptions.DoesNotExist:
                    pass  # Optionally log or handle invalid input
        feedback.is_active = True
        feedback.save()
        return render(request, 'thank_you.html', {'user': feedback.user})

    return render(request, 'feedback_response_form.html', {
        'feedback': feedback,
        'questions': questions,
        'options': feedback_options
    })


def thank_you(request):
    return render(request, 'thank_you.html')

from django.shortcuts import render
from .models import Feedback, FeedbackType, FeedbackResponse
from django.core.paginator import Paginator
from django.contrib.auth.decorators import login_required
from django.contrib.admin.views.decorators import user_passes_test




from collections import Counter
from django.shortcuts import render
from .models import FeedbackType, Feedback, FeedbackQuestion, FeedbackResponse, FeedbackOptions
from django.db.models import Avg
from .models import FeedbackType, FeedbackQuestion, FeedbackResponse
from django.db.models import Avg, Max
from .models import FeedbackType, FeedbackQuestion, FeedbackResponse, FeedbackOptions
@login_required
def analytical_feedback_view(request):
    feedback_data = []
    max_star = FeedbackOptions.objects.aggregate(max_val=Max('value'))['max_val'] or 5  # default to 5

    feedback_types = FeedbackType.objects.all()

    for f_type in feedback_types:
        total = Feedback.objects.filter(feedback_type=f_type).count()
        active = Feedback.objects.filter(feedback_type=f_type, is_active=True).count()
        percentage = (active / total) * 100 if total > 0 else 0

        questions = FeedbackQuestion.objects.filter(feedback_type=f_type)

        question_data = []
        for fq in questions:
            responses = FeedbackResponse.objects.filter(
                feedback__feedback_type=f_type,
                question=fq.question
            )
            avg_score = responses.aggregate(avg=Avg('options__value'))['avg'] or 0
            question_data.append({
                'question': fq.question.question,
                'avg_score': round(avg_score, 1),
                'id': fq.question.id, 
            })

        feedback_data.append({
            'type': f_type.type,
            'total': total,
            'active': active,
            'percentage': percentage,
            'questions': question_data
        })

    context = {
        'feedback_data': feedback_data,
        'max_star': max_star
    }
    return render(request, 'analytical_feedback.html', context)

from django.shortcuts import render, get_object_or_404
from .models import Question, FeedbackResponse, Feedback, FeedbackOptions
@login_required
def question_detail_view(request, question_id):
    question = get_object_or_404(Question, pk=question_id)

    # Get all Feedback assigned that includes this question (via FeedbackType)
    feedbacks_with_question = Feedback.objects.filter(
        feedback_type__feedbackquestion__question=question
    ).distinct()

    responded_feedback_ids = FeedbackResponse.objects.filter(
        question=question
    ).values_list('feedback_id', flat=True)

    non_responded_feedbacks = feedbacks_with_question.exclude(id__in=responded_feedback_ids)

    # Group responses by option
    options = FeedbackOptions.objects.all()
    responses_grouped = {}
    for option in options:
        users = FeedbackResponse.objects.filter(
            question=question,
            options=option
        ).select_related('feedback__user')
        if users.exists():
            responses_grouped[option.options] = [
                {
                    'username': r.feedback.user.username,
                    'first_name': r.feedback.user.first_name,
                    'last_name': r.feedback.user.last_name,
                    'submitted_on': r.submitted_on
                }
                for r in users
            ]

    context = {
        'question': question,
        'responses_grouped': responses_grouped,
        'non_responders': [
            {
                'username': f.user.username,
                'first_name': f.user.first_name,
                'last_name': f.user.last_name,
                'assigned_on': f.assigned_on
            }
            for f in non_responded_feedbacks
        ],
    }
    return render(request, 'feedback_question_detail.html', context)

from django.http import JsonResponse
from django.contrib.auth.decorators import login_required
from django.db.models import Count
from .models import Feedback, FeedbackType
from django.contrib.auth.models import User

@login_required
def feedback_summary_api(request, user_id):
    try:
        author_user = User.objects.get(pk=user_id)
    except User.DoesNotExist:
        return JsonResponse({'error': 'User not found'}, status=404)

    feedback_types = FeedbackType.objects.all()

    summary = []
    for ftype in feedback_types:
        sent_count = Feedback.objects.filter(
            feedback_type=ftype,
            user=author_user
        ).count()
        
        received_count = Feedback.objects.filter(
            feedback_type=ftype,
            user=author_user,
            is_active=True
        ).count()

        summary.append({
            'type': ftype.type,
            'sent_count': sent_count,
            'received_count': received_count,
        })

    return JsonResponse({'summary': summary})

# all feedbacks
from django.core.paginator import Paginator, EmptyPage
from django.db.models import Q
from django.shortcuts import render
from .models import Feedback, FeedbackType
from django.contrib.auth.models import User

@login_required
@user_passes_test(lambda u: is_eic(u) or is_super_admin(u))
def feedback_list(request):
    feedback_types = FeedbackType.objects.all()
    selected_type = request.GET.get('feedback_type')
    search_query = request.GET.get('search', '')
    status_filter = request.GET.get('status')
    print("status filter",status_filter)
    feedbacks = Feedback.objects.select_related('user', 'feedback_type')

    if status_filter == 'received':
        feedbacks = feedbacks.filter(is_active=True)
    elif status_filter == 'pending':
        feedbacks = feedbacks.filter(is_active=False)

    if selected_type:
        feedbacks = feedbacks.filter(feedback_type__id=selected_type)

    if search_query:
        feedbacks = feedbacks.filter(
            Q(user__first_name__icontains=search_query) |
            Q(user__last_name__icontains=search_query)
        )

    # Order by assigned_on (newest first)
    feedbacks = feedbacks.order_by('-assigned_on')

    # Pagination
    paginator = Paginator(feedbacks, 10)
    page_number = request.GET.get('page')
    try:
        page_obj = paginator.get_page(page_number)
    except EmptyPage:
        page_obj = paginator.get_page(1)

    # Totals
    total_sent = feedbacks.count()
    total_received = feedbacks.filter(is_active=True).count()

    context = {
        'data': page_obj,
        'feedback_types': feedback_types,
        'selected_type': selected_type,
        'status_filter': status_filter,
        'search_query': search_query,
        'total_sent': total_sent,
        'total_received': total_received
    }

    return render(request, 'feedback_list.html', context)
