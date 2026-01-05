from django.urls import path
from . import views


urlpatterns = [
path("", views.course_query, name="course_query"),


path("profile/", views.profile_view, name="profile"),
path("logout/", views.logout_view, name="logout"),


path("import/", views.import_excel, name="import_excel"),
path("import_all/", views.import_all_excels, name="import_all_excels"),


path("add-course/", views.add_course, name="add_course"),
path("delete-course/<int:course_id>/", views.delete_course, name="delete_course"),


# ✅ 個人課表（session 版）
path("personal/add/<int:course_id>/", views.add_personal_course, name="add_personal_course"),
path("personal/remove/<int:course_id>/", views.remove_personal_course, name="remove_personal_course"),
path("demo-logout/", views.demo_login_view, name="demo_logout"),
path("teacher-info/", views.teacher_info, name="teacher_info"),
path("debug_db/", views.debug_db, name="debug_db"),
path("demo-login/", views.demo_login_view, name="demo_login_view"),
]
