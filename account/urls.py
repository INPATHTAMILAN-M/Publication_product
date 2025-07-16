# urls.py
from django.urls import path
from .views import *
from django.contrib.auth import views as auth_views 
from django.contrib.auth.views import PasswordChangeDoneView
from oss.views import Draftview
from dl.views import bridge
from .whatsapp import send_message_view
from .sms import send_sms_view

urlpatterns = [
    path('register/',register, name='register'),
    path('login/', login_view, name='login'),
    path('user_management/', user_management, name='user_management'),
    path('publication_management/', publication_management, name='publication_management'),
    path('date_settings/', date_settings, name='date_settings'),
    path('modes/', modes, name='modes'),
    path('search-users/', search_users, name='search_users'),
    path('create_user/', create_user, name='create_user'),
    path('get_journals/', get_journals, name='get_journals'),
    path('assign_journal/', assign_journal, name='assign_journal'),
    path('fetch-groups/', fetch_groups, name='fetch_groups'),
    path('update-user-groups/', update_user_groups, name='update_user_groups'),
    
    path('fetch-reviewer-specializations/', fetch_reviewer_specializations, name='fetch_reviewer_specializations'),
    path('add_specialization/', add_specialization, name='add_specialization'),
    # path('fetch_specializations/',fetch_specializations, name='fetch_specializations'),
    
    path('reset_user_password/', reset_user_password, name='reset_user_password'),
    path('verify_email/<uidb64>/<token>/', verify_email, name='verify_email'),
    path('set_new_password/<uidb64>/<token>/', set_new_password, name='set_new_password'),
    path('registration_complete/',registration_complete, name='registration_complete'),
    path('index/', index, name='index'),
    path('', home, name='home'),
    path('base/', base_view, name='base'),   # URL for a general page
    path('password_reset/', password_reset_view, name='password_reset'),
    path('password_reset/done/', auth_views.PasswordResetDoneView.as_view(template_name='password_reset_done.html'), name='password_reset_done'),
    path('reset/<uidb64>/<token>/', CustomPasswordResetConfirmView.as_view(template_name='password_reset_confirm.html'), name='password_reset_confirm'),
    path('reset/done/', auth_views.PasswordResetCompleteView.as_view(template_name='password_reset_complete.html'), name='password_reset_complete'),
    path('user-management/',user_management, name='user_management'),
    path('assign-journal/', assign_journal, name='assign_journal'),
    path('change_password/', CustomPasswordChangeView.as_view(), name='change_password'),
    path('password_change_done/', PasswordChangeDoneView.as_view(template_name='password_change_done.html'), name='password_change_done'),
    path('profile/', profile, name='profile'),
    path('remove_specialization/<int:spec_id>/', remove_specialization, name='remove_specialization'),
    path('available-specializations/', available_specializations, name='available_specializations'),
    path('add-specializations/', add_specializations, name='add_specializations'),
    path('specializations/create/', create_specialization, name='create_specialization'),
    path('editor_profile/', editor_profile, name='editor_profile'),
    
    #logout
    path('logout/',custom_logout, name='logout'),
    
    #login redirect
    path('draft/', Draftview, name='draft'),
    #dl
    path('bridge/', bridge, name='bridge'),
    # ------------------Base for online paper submission-------------------------------------------------------------------------------------------------------------------------------------------------
    path('check_user_status/', check_user_status, name='check_user_status'),
    #current issues
    path('journal/<int:journal_id>/', current_issue_view, name='current_issue'),
    # dl related
    path('journal/<int:journal_id>/archives/',archive_view, name='archives'),
    path('journal/<int:journal_id>/volume/<int:volume_id>/issue/<int:issue_id>/',issue_detail_view, name='issue_detail'),
    # date setting
    # Add JOurnal
    path('add_journal/', add_journal, name='add_journal'),
    path('download/<int:article_id>/', download_article, name='download_article'),
    
    # path('upload-docx/', convert_docx_to_pdf, name='convert_docx'),
    # path('upload/', upload_file, name='upload'),
    path('compile-pdf/', compile_pdf, name='compile_pdf'),
    
    path('send_whatsapp/', send_message_view, name='send_whatsapp'),
    path('send_sms/', send_sms_view, name='send_sms'),
    
    # Feedback related URLs
    path('feedback_types/', feedback_types, name='feedback_types'),
    path('feedback_questions/<int:feedback_type_id>/', feedback_questions, name='feedback_questions'),
    path('remove_feedback_question/', remove_feedback_question, name='remove_feedback_question'),
    path('add_question_to_type/', add_question_to_type, name='add_question_to_type'),
    path('create_feedback_type/', create_feedback_type, name='create_feedback_type'),
    path('update_question_for_feedback_type/', update_question_for_feedback_type, name='update_question_for_feedback_type'),
    path('create_feedback/', create_feedback, name='create_feedback'),
    path('feedback/<int:feedback_id>/respond/', feedback_response_form, name='feedback_response_form'),
    path('feedback_list/', feedback_list, name='feedback_list'),
    path('analytical_feedback/', analytical_feedback_view, name='analytical_feedback'),
    path('feedback-summary/<int:user_id>/', feedback_summary_api, name='feedback_summary'),
    # path('feedback_list/',feedback_list, name='feedback_summary'),
    path('question/<int:question_id>/detail/', question_detail_view, name='question_detail'),
    path('thank-you/', thank_you, name='thank_you'),
]