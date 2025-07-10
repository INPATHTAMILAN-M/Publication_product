from django.urls import path
from .views import *
from django.conf import settings
from django.conf.urls.static import static
from account.views import custom_logout

urlpatterns = [
    # ------------------------Author submission ----------------------------------------------------------------------------------------------------------------------
    
    #Author Dashboard
    path('startnew/',startnew, name='startnew'),
    path('upload-additional-file/', upload_additional_file, name='upload_additional_file'),
    path('draft/',Draftview, name='draft'),
    path('submitted/',Submittedview, name='submitted'),
    path('revision/',Revisionview, name='revision'),
    path('accepted/',Acceptedview, name='accepted'),
    path('correction-comments/<int:submission_id>/', correction_comments_api, name='correction_comments_api'),
    path('upload-additional-corrected-file/', upload_additional_corrected_file, name='upload_additional_corrected_file'),
    path('upload_copyright_form/<int:submission_id>/',upload_copyright_form, name='upload_copyright_form'),
    path('rejected/',Rejectedview, name='rejected'),
    #Submission flow
    path('new_submission/<int:submission_id>/', submission_step_one, name='new_submission'),
    path('submission_step_two/<int:submission_id>/',submission_step_two, name='submission_step_two'),
    path('step3/<int:submission_id>/', keyword, name='step3'),
    path('step4/<int:submission_id>/', add_authors_institutions, name='step4'),
    path('step5/<int:submission_id>/', step5, name='step5'),
    path('step6/<int:submission_id>/', step6_review_submit, name='step6_review_submit'),
    #File processsing
    path('submission/<int:submission_id>/view_proof/',view_proof, name='view_proof'),
    path('submission_step_six/<int:submission_id>/', submission_step_six, name='submission_step_six'),
    #test
    #logout
    path('logout/', custom_logout, name='logout'),
    #ajax
    path('add-coauthor-ajax/<int:submission_id>/', add_coauthor_ajax, name='add_coauthor_ajax'),
    path('remove-coauthor-ajax/<int:submission_id>/', remove_coauthor_ajax, name='remove_coauthor_ajax'),


    # ------------------------Admin Office----------------------------------------------------------------------------------------------------------------
    path('admin_office/', admin_office_view, name='admin_office'),
    path('move_to_eic/<int:submission_id>/', move_to_eic, name='move_to_eic'),
    path('upload_plag_report/', upload_plag_report, name='upload_plag_report'),
    path('remove_plag_report/<int:submission_id>/', remove_plag_report, name='remove_plag_report'),
    path('reject_manuscript/<int:manuscript_id>/', reject_manuscript, name='reject_manuscript'),
    path('assign_ae/<int:submission_id>/', assign_ae, name='assign_ae'),
    
    # manuscripts rejection
    path('manuscripts_rejection/',Manuscripts_rejection, name='manuscripts_rejection'),
    
    # manuscripts acceptance
    path('manuscripts_acceptance',Manuscripts_acceptance, name='manuscripts_acceptance'),
    path('send_correction_report/', send_correction_report, name='send_correction_report'),
    path('get-correction-file/', get_correction_file, name='get_correction_file'),
    path('upload_corrected_file/',upload_corrected_file, name='upload_corrected_file'),
    
    # manuscripts in review
    path('manuscripts_review/',manuscripts_review, name='manuscripts_review'),
    
    # Manuscripts in revision
    path('manuscripts_revision_overdue/',Manuscripts_revision_overdue, name='manuscripts_revision_overdue'),
    path('manuscripts_revision/',Manuscripts_revision, name='manuscripts_revision'),
    
    # Manuscripts in proof
    path('setting_proof/',Setting_proof, name='setting_proof'),
    path('submission_details/<int:submission_id>/', submission_details, name='submission_details'),
    path('reject_submission/', reject_submission, name='reject_submission'),
    path('get_submission_details/', get_submission_details, name='get_submission_details'),
    path('upload_typeset_document/', upload_typeset_document, name='upload_typeset_document'),
    path('typeset/<int:submission_id>/', view_typeset_document, name='view_typeset_document'),
    path('mark_proof_read_done/<int:submission_id>/', mark_proof_read_done, name='mark_proof_read_done'),
    path('history/', history, name='history'),
    # -------------------------Editor In hief-------------------------------------------------------------------------------------------------------------
    path('eic_review_manuscripts/', eic_review_manuscripts, name='eic_review_manuscripts'),
    path('editor_in_chief/',Editor_in_Chief, name='editor_in_chief'),
    path('save_decision/', save_decision, name='save_decision'),
    path('decisioned_manuscripts/', decisioned_manuscripts, name='decisioned_manuscripts'),
    # -------------------------Reviewer-------------------------------------------------------------------------------------------------------------------
    path('reviewer_invitations', reviewer_invitations, name='reviewer_invitations'),
    path('accept_invitation', accept_invitation, name='accept_invitation'),
    path('reject_invitation', reject_invitation, name='reject_invitation'),
    path('manuscripts_to_review/', Manuscripts_To_Review, name='manuscripts_to_review'),
    path('reviewed_manuscripts/', Reviewed_Manuscripts, name='reviewed_manuscripts'),
    path('submit_review_comments/',submit_review_comments, name='submit_review_comments'),
    path('get_reviewer_details/',get_reviewer_details, name='get_reviewer_details'),

    # -------------------------Associate Editor--------------------------------------------------------------------------------------------------------------------
    path('associate_editor/', associate_editor, name='associate_editor'),
    path('get_reviewers_comments/', get_reviewers_comments, name='get_reviewers_comments'),
    path('submit_recommendation/', submit_recommendation, name='submit_recommendation'),
    # path('get_reviewer_details/', get_reviewer_details, name='get_reviewer_details'),
    path('manuscripts_under_review/', manuscripts_under_review, name='manuscripts_under_review'),
    path('manuscripts_eic/', manuscripts_eic, name='manuscripts_eic'),
    path('get_reviewers/', get_reviewers, name='get_reviewers'),
    path('send_invitation/', send_invitation, name='send_invitation'),
    path('cancel_invitation/', cancel_invitation, name='cancel_invitation'),
    path('manuscripts_review_report/',Manuscripts_Review_Report, name='manuscripts_review_report'),
    
    # success for mail
    path('contact_form/', contact, name='contact_form'),
    path('success/', success_page, name='success_page'),
    
    path('export_submissions/', export_submissions_to_excel, name='export_submissions'),

    # dashboard
    path('dashboard/', dashboard_view, name='dashboard'),

]+ static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
 