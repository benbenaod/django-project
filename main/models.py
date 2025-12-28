from django.db import models
from django.conf import settings


class Teacher(models.Model):
    user = models.OneToOneField(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="teacher_profile",
    )

    TYPE_CHOICES = [
        ("專任", "專任"),
        ("兼任", "兼任"),
        ("其他", "其他"),
    ]

    name_ch = models.CharField("中文姓名", max_length=50)
    name_en = models.CharField("英文姓名", max_length=100, blank=True, default="")
    extension = models.CharField("校內分機", max_length=20, blank=True, default="")
    email = models.EmailField("E-mail", blank=True, default="")
    office = models.CharField("辦公室/地點", max_length=100, blank=True, default="")
    intro = models.TextField("老師簡介", blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name_ch


class Student(models.Model):
    user = models.OneToOneField(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="student_profile",
    )

    student_id = models.CharField("學號", max_length=20, unique=True)
    name = models.CharField("姓名", max_length=50)
    email = models.EmailField("E-mail", blank=True, default="")
    department_code = models.CharField("系所代碼", max_length=20, blank=True, default="")
    grade = models.CharField("年級", max_length=10, blank=True, default="")
    class_group = models.CharField("班組", max_length=20, blank=True, default="")
    is_active = models.BooleanField("啟用", default=True)
    created_at = models.DateTimeField(auto_now_add=True, null=True, blank=True)

    def __str__(self):
        return f"{self.student_id} - {self.name}"


class Course(models.Model):
    # 編號/學期/教師
    number = models.CharField(max_length=10, blank=True, default="")
    semester = models.CharField(max_length=10)
    teacher = models.CharField(max_length=50, blank=True, default="")

    # 課程代碼/系所
    course_code = models.CharField(max_length=50, blank=True, default="")
    department_code = models.CharField(max_length=20, blank=True, default="")
    core_code = models.CharField(max_length=20, blank=True, default="")
    group_code = models.CharField(max_length=20, blank=True, default="")

    # 查詢用欄位
    grade = models.CharField(max_length=10, blank=True, default="")
    class_group = models.CharField(max_length=20, blank=True, default="")
    course_name = models.CharField(max_length=200)
    division = models.CharField(max_length=20, blank=True, default="")  # 課別名稱
    system = models.CharField(max_length=30, blank=True, default="")    # 學制別

    teaching_group = models.CharField(max_length=50, blank=True, default="")
    week_info = models.CharField(max_length=50, blank=True, default="")
    day = models.CharField(max_length=5, blank=True, default="")
    period = models.CharField(max_length=20, blank=True, default="")
    classroom = models.CharField(max_length=100, blank=True, default="")

    course_summary_ch = models.TextField(null=True, blank=True)
    course_summary_en = models.TextField(null=True, blank=True)

    teacher_old_code = models.CharField(max_length=20, blank=True, default="")
    course_old_code = models.CharField(max_length=20, blank=True, default="")
    schedule_old_code = models.CharField(max_length=20, blank=True, default="")
    schedule_old_name = models.CharField(max_length=50, blank=True, default="")
    teacher_old_code2 = models.CharField(max_length=20, blank=True, default="")

    teacher_ref = models.ForeignKey(
        Teacher,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="courses",
        verbose_name="老師(關聯)",
    )

    def __str__(self):
        return f"{self.course_name} - {self.teacher}"


class Enrollment(models.Model):
    student = models.ForeignKey(Student, on_delete=models.CASCADE, related_name="enrollments")
    course = models.ForeignKey(Course, on_delete=models.CASCADE, related_name="enrollments")
    created_at = models.DateTimeField(auto_now_add=True, null=True, blank=True)
    class Meta:
        unique_together = ("student", "course")

    def __str__(self):
        return f"{self.student.student_id} - {self.course.course_name}"
