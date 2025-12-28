from django import forms
from .models import Course

class CourseForm(forms.ModelForm):
    # 星期下拉
    DAY_CHOICES = [
        ("1", "星期一"),
        ("2", "星期二"),
        ("3", "星期三"),
        ("4", "星期四"),
        ("5", "星期五"),
        ("6", "星期六"),
        ("7", "星期日"),
    ]

    day = forms.ChoiceField(choices=DAY_CHOICES, label="星期")

    class Meta:
        model = Course
        fields = [
            "department_code",  # 系所
            "grade",            # 年級
            "class_group",      # ✅ 班組（你要新增的）
            "division",         # 課別
            "course_name",      # 科目名稱
            "classroom",        # 教室
            "day",              # 星期
            "period",           # 節次
        ]
        labels = {
            "department_code": "系所",
            "grade": "年級",
            "class_group": "班組",
            "division": "課別",
            "course_name": "科目名稱",
            "classroom": "教室",
            "period": "節次",
        }
        widgets = {
            "department_code": forms.TextInput(attrs={"class": "text-input", "placeholder": "例如 22140 / 22160"}),
            "grade": forms.TextInput(attrs={"class": "text-input", "placeholder": "例如 1 / 2 / 3 / 4"}),
            "class_group": forms.TextInput(attrs={"class": "text-input", "placeholder": "例如 A0 / B0 / 13"}),
            "division": forms.TextInput(attrs={"class": "text-input", "placeholder": "例如 專業必修(系所)"}),
            "course_name": forms.TextInput(attrs={"class": "text-input", "placeholder": "例如 系統分析與設計"}),
            "classroom": forms.TextInput(attrs={"class": "text-input", "placeholder": "例如 F602"}),
            "period": forms.TextInput(attrs={"class": "text-input", "placeholder": "例如 2,3,4 或 8-10"}),
            "day": forms.Select(attrs={"class": "text-input"}),  # 讓下拉跟你版型一致
        }
