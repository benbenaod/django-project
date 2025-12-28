from pathlib import Path
from django.views.decorators.http import require_GET
import pandas as pd
from django.contrib import messages
from django.contrib.auth import (
    authenticate,
    get_user_model,
    login,
    logout,
    update_session_auth_hash,
)
from django.db.models import Q
from django.http import HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.utils.html import escape
from django.views.decorators.http import require_POST
import json
from .forms import CourseForm
from .models import Course, Student, Teacher

BUILDING_URL_MAP = {
    "F": "https://www.google.com/maps/place/%E5%9C%8B%E7%AB%8B%E8%87%BA%E5%8C%97%E8%AD%B7%E7%90%86%E5%81%A5%E5%BA%B7%E5%A4%A7%E5%AD%B8%E5%AD%B8%E6%80%9D%E6%A8%93/@25.1186186,121.5166288,17z/data=!3m1!4b1!4m6!3m5!1s0x3442af4ac9da7987:0xf36d626d63834f5!8m2!3d25.1186138!4d121.5192037!16s%2Fg%2F11s82z2lrp?entry=ttu&g_ep=EgoyMDI1MTIwOS4wIKXMDSoKLDEwMDc5MjA2OUgBUAM%3D",
    "S": "https://www.google.com/maps/place/%E5%9C%8B%E7%AB%8B%E8%87%BA%E5%8C%97%E8%AD%B7%E7%90%86%E5%81%A5%E5%BA%B7%E5%A4%A7%E5%AD%B8%E7%A7%91%E6%8A%80%E5%A4%A7%E6%A8%93/@25.117542,121.5180909,17z/data=!3m1!5s0x3442ae8a4f198def:0x16fcf46afefac4c2!4m16!1m9!3m8!1s0x3442ae8967e29825:0xa74a929b7ae3dbf6!2z5ZyL56uL6Ie65YyX6K2355CG5YGl5bq35aSn5a2456eR5oqA5aSn5qiT!8m2!3d25.1175372!4d121.5206658!9m1!1b1!16s%2Fg%2F11b6jgqh03!3m5!1s0x3442ae8967e29825:0xa74a929b7ae3dbf6!8m2!3d25.1175372!4d121.5206658!16s%2Fg%2F11b6jgqh03?entry=ttu&g_ep=EgoyMDI1MTIwOS4wIKXMDSoKLDEwMDc5MjA2OUgBUAM%3D",
    "B": "https://www.google.com/maps/place/%E5%9C%8B%E7%AB%8B%E8%87%BA%E5%8C%97%E8%AD%B7%E7%90%86%E5%81%A5%E5%BA%B7%E5%A4%A7%E5%AD%B8%E8%A6%AA%E4%BB%81%E6%A8%93/@25.1185795,121.5185797,17z/data=!3m2!4b1!5s0x3442ae8a4f198def:0x16fcf46afefac4c2!4m6!3m5!1s0x3442af851c386faf:0xc3edb631a5715fd3!8m2!3d25.1185747!4d121.5211546!16s%2Fg%2F11ryljg7x2?entry=ttu&g_ep=EgoyMDI1MTIwOS4wIKXMDSoKLDEwMDc5MjA2OUgBUAM%3D",
    "G": "https://www.google.com.tw/maps/place/%E5%9C%8B%E7%AB%8B%E8%87%BA%E5%8C%97%E8%AD%B7%E7%90%86%E5%81%A5%E5%BA%B7%E5%A4%A7%E5%AD%B8%E6%A0%A1%E6%9C%AC%E9%83%A8/@25.1175841,121.5166108,17z/data=!3m2!4b1!5s0x3442ae8a4f198def:0x16fcf46afefac4c2!4m6!3m5!1s0x3442ae8bc54ebc79:0xfd2a9d659e97b078!8m2!3d25.1175793!4d121.5214817!16s%2Fm%2F0z8mtpb?entry=ttu&g_ep=EgoyMDI1MTIwOS4wIKXMDSoASAFQAw%3D%3D",
}

# ==============================
# ✅ 預設帳號/密碼 + 建立 Teacher/Student profile
# ==============================
_DEFAULT_CREATED = False


def ensure_default_accounts():
    """在每次進入 course_query 時確保預設帳號存在，並綁定 Teacher/Student。"""
    global _DEFAULT_CREATED
    if _DEFAULT_CREATED:
        return

    User = get_user_model()

    DEFAULT_ACCOUNTS = [
        {"role": "teacher", "username": "dora", "password": "a", "teacher_name": "中岳"},
        {
            "role": "student",
            "username": "ben",
            "password": "a",
            "student_id": "122214132",
            "student_name": "童國原",
        },
    ]

    for item in DEFAULT_ACCOUNTS:
        username = (item.get("username") or "").strip()
        password = item.get("password") or ""
        if not username:
            continue

        user, created = User.objects.get_or_create(username=username)
        if created:
            user.set_password(password)
            user.save()

        if item["role"] == "teacher":
            teacher_name = (item.get("teacher_name") or username).strip()

            if not (user.first_name or "").strip():
                user.first_name = teacher_name
                user.save()

            t = Teacher.objects.filter(name_ch=teacher_name, user__isnull=True).first()
            if t:
                t.user = user
                t.save()
            else:
                Teacher.objects.get_or_create(user=user, defaults={"name_ch": teacher_name})

        elif item["role"] == "student":
            sid = (item.get("student_id") or username).strip()
            sname = (item.get("student_name") or username).strip()

            if not (user.first_name or "").strip():
                user.first_name = sname
                user.save()

            s = Student.objects.filter(student_id=sid).first()
            if s:
                if getattr(s, "user_id", None) is None:
                    s.user = user
                    s.save()
            else:
                Student.objects.get_or_create(
                    user=user,
                    defaults={"student_id": sid},
                )

    _DEFAULT_CREATED = True


def profile_view(request):
    """處理個人資料管理彈窗送出的『更新密碼』"""
    if request.method == "POST":
        new_password = (request.POST.get("new_password") or "").strip()
        confirm_password = (request.POST.get("confirm_password") or "").strip()

        if not new_password:
            messages.error(request, "新密碼不能為空白。")
        elif new_password != confirm_password:
            messages.error(request, "新密碼與確認密碼不一致。")
        else:
            user = request.user
            user.set_password(new_password)
            user.save()
            update_session_auth_hash(request, user)
            messages.success(request, "密碼已更新，下一次登入請使用新密碼。")

        return redirect("course_query")

    return redirect("course_query")


def logout_view(request):
    """登出：只接受 POST，比較安全"""
    if request.method == "POST":
        logout(request)
    return redirect("course_query")


# ==============================
#   找 Excel 資料夾（MyDrive / My Drive 兩種情況）
# ==============================


def get_excel_dir():
    candidates = [
        Path("/content/drive/MyDrive/python/系統分析/課程查詢"),
        Path("/content/drive/My Drive/python/系統分析/課程查詢"),
    ]
    for p in candidates:
        if p.exists():
            print(f"✅ 使用 Excel 資料夾：{p}")
            return p
    print("⚠️ 找不到任何有效的 Excel 資料夾，請確認路徑。")
    return candidates[0]


EXCEL_DIR = get_excel_dir()


# ==============================
#   小工具：任何值 → 安全字串（處理 NaN/None/"nan"）
# ==============================


def safe_str(v):
    if v is None:
        return ""
    try:
        import pandas as _pd

        if _pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def esc(v):
    """安全輸出到 HTML attribute / text"""
    return escape(safe_str(v))


def safe_get(row, col_name, default=""):
    try:
        return safe_str(row.get(col_name, default))
    except Exception:
        return safe_str(default)


# ==============================
# ✅ 教室欄位：統一顯示/匯入
# ==============================

ROOM_COL_CANDIDATES = [
    "上課地點",
    "教室地點",
    "上課教室",
    "教室",
    "上課位置",
    "上課地點(教室)",
    "上課教室地點",
    "上課地點/教室",
    "地點",
    "位置",
]


def room_from_row(row) -> str:
    for col in ROOM_COL_CANDIDATES:
        v = safe_get(row, col)
        if v:
            return v
    return ""


def room_display(c: Course) -> str:
    v = safe_str(getattr(c, "classroom", ""))
    v = v or safe_str(getattr(c, "room", ""))
    v = v or safe_str(getattr(c, "location", ""))
    return v or "-"


# ==============================
# ✅ Teacher 取用：中文姓名 / 類別 / 分機
# ==============================


def _teacher_meta_from_obj(t: Teacher):
    if not t:
        return "", "", ""

    name_ch = safe_str(getattr(t, "name_ch", ""))

    category = (
        safe_str(getattr(t, "category", ""))
        or safe_str(getattr(t, "type", ""))
        or safe_str(getattr(t, "role", ""))
        or safe_str(getattr(t, "title", ""))
    )

    ext = (
        safe_str(getattr(t, "extension", ""))
        or safe_str(getattr(t, "ext", ""))
        or safe_str(getattr(t, "phone_ext", ""))
        or safe_str(getattr(t, "school_ext", ""))
    )

    return name_ch, category, ext


def teacher_meta_for_course(c: Course):
    if not c:
        return "", "", ""

    t = getattr(c, "teacher_ref", None)

    if not t:
        tname = safe_str(getattr(c, "teacher", ""))
        if tname:
            t = Teacher.objects.filter(name_ch=tname).first()

    name_ch, category, ext = _teacher_meta_from_obj(t)

    if not name_ch:
        name_ch = safe_str(getattr(c, "teacher", ""))

    return name_ch, category, ext


# ==============================
#   DataFrame → Course 資料表（含 Teacher 自動對應）
# ==============================


def _import_df_to_course(df: pd.DataFrame) -> int:
    if "科目中文名稱" not in df.columns:
        print("⚠️ Excel 裡找不到『科目中文名稱』欄位，請確認欄位名稱。")
        print("目前欄位：", list(df.columns))
        return 0

    df = df.dropna(subset=["科目中文名稱"])
    count = 0
    teacher_cache = {}

    for _, row in df.iterrows():
        course_name = safe_get(row, "科目中文名稱")
        if not course_name:
            continue

        teacher_name = safe_get(row, "主開課教師姓名")
        teacher_obj = None

        if teacher_name:
            teacher_obj = teacher_cache.get(teacher_name)
            if teacher_obj is None:
                teacher_obj, _ = Teacher.objects.get_or_create(
                    name_ch=teacher_name,
                    defaults={"name_en": ""},
                )
                teacher_cache[teacher_name] = teacher_obj

        classroom_val = room_from_row(row)

        Course.objects.create(
            number=safe_get(row, "編號"),
            semester=safe_get(row, "學期"),
            teacher=teacher_name,
            course_code=safe_get(row, "科目代碼(新碼全碼)"),
            department_code=safe_get(row, "系所代碼"),
            core_code=safe_get(row, "核心四碼"),
            group_code=safe_get(row, "科目組別"),
            grade=safe_get(row, "年級"),
            class_group=safe_get(row, "上課班組"),
            course_name=course_name,
            division=safe_get(row, "課別名稱"),
            system=safe_get(row, "學制別"),
            teaching_group=safe_get(row, "授課群組"),
            week_info=safe_get(row, "上課週次"),
            day=safe_get(row, "上課星期"),
            period=safe_get(row, "上課節次"),
            classroom=classroom_val,
            course_summary_ch=safe_get(row, "課程中文摘要"),
            course_summary_en=safe_get(row, "課程英文摘要"),
            teacher_old_code=safe_get(row, "主開課教師代碼(舊碼)"),
            course_old_code=safe_get(row, "科目代碼(舊碼)"),
            schedule_old_code=safe_get(row, "課表代碼(舊碼)"),
            schedule_old_name=safe_get(row, "課表名稱(舊碼)"),
            teacher_old_code2=safe_get(row, "授課教師代碼(舊碼)"),
            teacher_ref=teacher_obj,
        )
        count += 1

    return count


def ensure_courses_loaded():
    if Course.objects.exists():
        return

    excel_files = sorted(EXCEL_DIR.glob("*.xlsx"))
    if not excel_files:
        print(f"⚠️ 在 {EXCEL_DIR} 裡沒有找到任何 .xlsx 檔案")
        return

    print(f"🔄 資料表為空，開始自動匯入 Excel（共 {len(excel_files)} 個檔案）...")
    for file_path in excel_files:
        try:
            print(f"➡ 讀取 {file_path}")
            df = pd.read_excel(file_path, header=4)
            _import_df_to_course(df)
        except Exception as e:
            print(f"❌ 讀取 {file_path} 失敗：{e}")


# ==============================
# ✅ 穩定抓使用者顯示姓名
# ==============================


def get_user_display_name(user):
    if not user or not getattr(user, "is_authenticated", False):
        return "", ""

    t = Teacher.objects.filter(user=user).first()
    if t:
        name = (
            safe_str(getattr(t, "name_ch", ""))
            or safe_str(getattr(user, "first_name", ""))
            or safe_str(getattr(user, "username", ""))
        )
        return "老師", name

    s = Student.objects.filter(user=user).first()
    if s:
        name = (
            safe_str(getattr(s, "name", ""))
            or safe_str(getattr(user, "first_name", ""))
            or safe_str(getattr(user, "username", ""))
        )
        return "學生", name

    first = safe_str(getattr(user, "first_name", ""))
    if first:
        return "使用者", first

    return "使用者", safe_str(getattr(user, "username", ""))


# ==============================
# ✅ 系所代碼 → 中文系所名（顯示用）
# ==============================
DEPT_NAME_MAP = {
    "22140": "資訊管理系",
    "22160": "資訊管理系碩士班",
    "11140": "護理系",
    "21140": "健康事業管理系",
    "24120": "二年制長期照護系",
    "13140": "高齡健康照護系",
    "31140": "嬰幼兒保育系",
    "25140": "語言治療與聽力學系",
    "23140": "休閒產業與健康促進系",
    "32140": "運動保健系",
    "33140": "生死與健康心理諮商系",
    "30860": "國際運動科學外國學生專班",
    "33161": "生死與健康心理諮商系碩士班生死學組",
    "33162": "生死與健康心理諮商系碩士班諮商心理組",
    "1C120": "二年制護理助產及婦女健康系",
    "1C160": "護理助產及婦女健康系護理助產碩士班",
    "1C330": "二年制進修部護理助產及婦女健康系",
    "1D120": "二年制醫護教育暨數位學習系",
    "1D160": "醫護教育暨數位學習系碩士班",
    "20160": "健康科技學院碩士班",
    "26860": "國際健康科技碩士學位學程國際生碩士班",
    "21120": "二年制健康事業管理系",
    "21160": "健康事業管理系碩士班",
    "21460": "健康事業管理系碩士在職專班",
    "21330": "二年制進修部健康事業管理系",
    "23160": "休閒產業與健康促進系旅遊健康碩士班",
    "23460": "休閒產業與健康促進系碩士在職專班",
    "24160": "長期照護系碩士班",
    "24150": "長期照護系學士後多元專長培力課程專班",
    "25161": "語言治療與聽力學系碩士班語言治療組",
    "25162": "語言治療與聽力學系碩士班聽力組",
    "25460": "語言治療與聽力學系在職專班",
    "31120": "二年制嬰幼兒保育系",
    "31160": "嬰幼兒保育系碩士班",
    "31860": "國際蒙特梭利碩士專班",
    "32160": "運動保健系碩士班",
    "32460": "運動保健系碩士在職專班",
    "11120": "二年制護理系",
    "11230": "二年制進修部護理系(日間班)",
    "11330": "二年制進修部護理系(夜間班)",
    "11860": "國際護理碩士班",
    "1C860": "國際護理助產碩士班",
    "43160": "人工智慧與健康大數據研究所",
    "32860": "國際運動科學暨智慧健康科技碩士專班",
    "42140": "智慧健康科技技優專班",
    "41140": "高齡與運動健康暨嬰幼兒保育技優專班",
    "11190": "護理系學士後學士班(學士後護理系)",
    "31180": "嬰幼兒保育系學士後教保學位學程",
    "11170": "護理系博士班",
    "11464": "護理系碩士在職專班護專精神組",
    "11462": "護理系碩士在職專班護專老人組",
    "11870": "國際護理博士班",
    "13160": "高齡健康照護系碩士班",
    "11161": "護理系碩士班護研成人組",
    "11169": "護理系中西醫結合護理碩士班",
    "11163": "護理系碩士班護研婦女組",
    "11165": "護理系碩士班護研社區組",
    "11167": "護理系碩士班護研資訊組",
    "11466": "護理系碩士在職專班護專兒童組",
    "11468": "護理系碩士在職專班護專成人專科組",
    "1D110": "二專醫護教育暨數位學習科",
    "11162": "護理系碩士班護研老人組",
    "11164": "護理系碩士班護研精神組",
    "11166": "護理系碩士班護研兒童組",
    "11168": "護理系碩士班護研成人專科組",
    "11461": "護理系碩士在職專班護專成人組",
    "11463": "護理系碩士在職專班護專婦女組",
    "11465": "護理系碩士在職專班護專社區組",
    "11467": "護理系碩士在職專班護專資訊組",
}


def dept_display(code: str) -> str:
    code = safe_str(code)
    return DEPT_NAME_MAP.get(code, "") or code or "-"


FOUR_TECH_DEPTS = {
    "22140": "資訊管理系",
    "22160": "資訊管理系碩士班",
    "11140": "護理系",
    "21140": "健康事業管理系",
    "24120": "二年制長期照護系",
    "13140": "高齡健康照護系",
    "31140": "嬰幼兒保育系",
    "25140": "語言治療與聽力學系",
    "23140": "休閒產業與健康促進系",
    "32140": "運動保健系",
    "33140": "生死與健康心理諮商系",
}


def apply_system_filter(qs, system_value: str):
    system_value = safe_str(system_value)
    if not system_value:
        return qs

    def hit(keyword: str) -> Q:
        keyword = safe_str(keyword)
        if not keyword:
            return Q()
        return (
            Q(system__icontains=keyword)
            | Q(schedule_old_name__icontains=keyword)
            | Q(schedule_old_code__icontains=keyword)
            | Q(class_group__icontains=keyword)
            | Q(teaching_group__icontains=keyword)
            | Q(course_name__icontains=keyword)
            | Q(department_code__icontains=keyword)
        )

    def dept_codes_by_keywords(keywords):
        codes = set()
        for code, name in DEPT_NAME_MAP.items():
            for k in keywords:
                if k and k in (name or ""):
                    codes.add(code)
                    break
        return codes

    if system_value == "二專":
        keywords = ["1D110"]
        codes = dept_codes_by_keywords(keywords)
        q = hit("1D110")
        if codes:
            q |= Q(department_code__in=list(codes))
        return qs.filter(q)

    if system_value == "二技":
        keywords = ["二年制"]
        codes = dept_codes_by_keywords(keywords)
        q = hit("二年制") | hit("二技")
        if codes:
            q |= Q(department_code__in=list(codes))
        return qs.filter(q)

    if system_value == "二技(三年)":
        keywords = ["二年制進修部"]
        codes = dept_codes_by_keywords(keywords)
        q = hit("二年制進修部") | hit("二技(三年)")
        if codes:
            q |= Q(department_code__in=list(codes))
        return qs.filter(q)

    if system_value == "四技":
        keywords = ["四年制", "四技"]
        codes = dept_codes_by_keywords(keywords) | set(FOUR_TECH_DEPTS.keys())
        q = hit("四年制") | hit("四技") | Q(department_code__in=list(codes))
        return qs.filter(q)

    if system_value == "學士後多元專長":
        keywords = ["學士後多元專長"]
        codes = dept_codes_by_keywords(keywords)
        q = hit("學士後多元專長")
        if codes:
            q |= Q(department_code__in=list(codes))
        return qs.filter(q)

    if system_value == "碩士班":
        keywords = ["研究所", "學生專班", "碩士班", "碩士在職"]
        codes = dept_codes_by_keywords(keywords)
        q = Q()
        for k in keywords:
            q |= hit(k)
        if codes:
            q |= Q(department_code__in=list(codes))
        return qs.filter(q)

    if system_value == "博士班":
        keywords = ["博士"]
        codes = dept_codes_by_keywords(keywords)
        q = hit("博士")
        if codes:
            q |= Q(department_code__in=list(codes))
        return qs.filter(q)

    if system_value == "學士後學位學程":
        keywords = ["學士後教保學位學程"]
        codes = dept_codes_by_keywords(keywords)
        q = hit("學士後教保學位學程")
        if codes:
            q |= Q(department_code__in=list(codes))
        return qs.filter(q)

    if system_value == "學士後系":
        keywords = ["學士後學士班"]
        codes = dept_codes_by_keywords(keywords)
        q = hit("學士後學士班")
        if codes:
            q |= Q(department_code__in=list(codes))
        return qs.filter(q)

    return qs


# ==============================
# ✅ 個人課表（Session 存 ids）
# ==============================
SESSION_KEY_PERSONAL = "personal_course_ids"

DEFAULT_PERSONAL_SEMESTER = "1141"
DEFAULT_PERSONAL_CLASS_GROUP = "A0"
REQUIRED_DEPT_FOR_RESEARCH = "22140"
REQUIRED_KEYWORDS = ["系統分析", "研究概論"]


def _get_personal_ids(request):
    ids = request.session.get(SESSION_KEY_PERSONAL, [])
    if not isinstance(ids, list):
        ids = []
    out = []
    for x in ids:
        try:
            xi = int(x)
            if xi not in out:
                out.append(xi)
        except Exception:
            continue
    return out


def _set_personal_ids(request, ids):
    request.session[SESSION_KEY_PERSONAL] = list(ids)
    request.session.modified = True


def get_required_personal_courses():
    base = Course.objects.filter(
        semester=DEFAULT_PERSONAL_SEMESTER,
        class_group__icontains=DEFAULT_PERSONAL_CLASS_GROUP,
    )

    rule = {
        "系統分析": {},
        "研究概論": {"department_code": REQUIRED_DEPT_FOR_RESEARCH},
    }
    return base, rule


def resolve_required_course_ids():
    base, rule = get_required_personal_courses()
    required_ids = {}
    for kw, extra in rule.items():
        qs = base.filter(course_name__icontains=kw)
        if extra.get("department_code"):
            qs = qs.filter(department_code__exact=extra["department_code"])
        c = qs.order_by("day", "period", "course_name").first()
        if c:
            required_ids[kw] = c.id
    return required_ids


def ensure_fixed_personal_courses(request):
    if not request.user.is_authenticated:
        return
    if not Student.objects.filter(user=request.user).exists():
        return

    required_map = resolve_required_course_ids()
    existing = _get_personal_ids(request)
    existing_set = set(existing)

    for kw in REQUIRED_KEYWORDS:
        rid = required_map.get(kw)
        if rid and rid not in existing_set:
            existing.append(rid)
            existing_set.add(rid)

    _set_personal_ids(request, existing)


def is_required_course_id(course_id: int) -> bool:
    req = resolve_required_course_ids()
    return course_id in set(req.values())


def required_remove_message(course_id: int) -> str:
    req = resolve_required_course_ids()
    inv = {v: k for k, v in req.items()}
    name = inv.get(course_id, "此課程")
    return f"【{name}】為必修安排，無法移除。"


def parse_periods(period_raw: str):
    raw = safe_str(period_raw)
    if not raw:
        return []
    raw = raw.replace("、", ",").replace(" ", "")
    out = []
    for part in raw.split(","):
        if not part:
            continue
        if "-" in part:
            try:
                a, b = part.split("-")
                a = int(a)
                b = int(b)
                for p in range(min(a, b), max(a, b) + 1):
                    if p not in out:
                        out.append(p)
            except Exception:
                continue
        else:
            try:
                p = int(part)
                if p not in out:
                    out.append(p)
            except Exception:
                continue
    return out


def _course_slots(course: Course):
    d = safe_str(getattr(course, "day", ""))
    if not d:
        return set()
    ps = parse_periods(safe_str(getattr(course, "period", "")))
    return {f"{d}-{p}" for p in ps}


def _conflict_slots(existing_courses, new_course: Course):
    exist_slots = set()
    for c in existing_courses:
        exist_slots |= _course_slots(c)
    new_slots = _course_slots(new_course)
    return sorted(list(exist_slots & new_slots))


def _format_conflicts(conflicts):
    day_map = {"1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "日"}
    items = []
    for k in conflicts:
        try:
            d, p = k.split("-")
        except Exception:
            continue
        items.append(f"星期{day_map.get(d, d)} 第{p}節")
    return "、".join(items)


def build_grid_timetable_html(courses, *, title: str):
    period_time_map = {
        1: "08:10~09:00",
        2: "09:10~10:00",
        3: "10:10~11:00",
        4: "11:10~12:00",
        5: "12:40~13:30",
        6: "13:40~14:30",
        7: "14:40~15:30",
        8: "15:40~16:30",
        9: "16:40~17:30",
        10: "17:40~18:30",
        11: "18:35~19:25",
        12: "19:30~20:20",
        13: "20:25~21:15",
        14: "21:20~22:10",
    }
    day_labels = [("1", "一"), ("2", "二"), ("3", "三"), ("4", "四"), ("5", "五"), ("6", "六"), ("7", "日")]
    period_range = list(range(1, 15))

    timetable = {}
    for c in courses:
        d = safe_str(getattr(c, "day", ""))
        per_raw = safe_str(getattr(c, "period", ""))
        if not d or not per_raw:
            continue
        for p in parse_periods(per_raw):
            timetable.setdefault(d, {}).setdefault(p, []).append(c)

    table_html = '<div class="timetable-wrapper">'
    table_html += f'<div class="timetable-title">{esc(title)}</div>'
    table_html += '<table class="timetable">'
    table_html += '<tr><th>節次</th>'
    for _val, label in day_labels:
        table_html += f"<th>星期{esc(label)}</th>"
    table_html += "</tr>"

    for p in period_range:
        t = period_time_map.get(p, "")
        th_html = f'{p}<div style="font-size:11px;color:#6b7280;margin-top:4px;">{esc(t)}</div>' if t else f"{p}"
        table_html += f"<tr><th>{th_html}</th>"

        for day_val, day_label in day_labels:
            cell_courses = timetable.get(day_val, {}).get(p, [])
            if not cell_courses:
                table_html += "<td>&nbsp;</td>"
                continue

            parts = []
            for c in cell_courses:
                week = safe_str(getattr(c, "week_info", ""))
                time_text = f"星期{day_label} 第{safe_str(getattr(c, 'period', ''))}節"
                if week:
                    time_text += f"（{week}）"

                t_ch, t_cat, t_ext = teacher_meta_for_course(c)
                room_txt = room_display(c)

                parts.append(
                    (
                        f'<div class="course-cell course-clickable" '
                        f'data-id="{c.id}" '
                        f'data-name="{esc(getattr(c, "course_name", ""))}" '
                        f'data-dept="{esc(dept_display(getattr(c, "department_code", "")))}" '
                        f'data-teacher="{esc(getattr(c, "teacher", ""))}" '
                        f'data-teacher-ch="{esc(t_ch)}" '
                        f'data-teacher-category="{esc(t_cat)}" '
                        f'data-teacher-ext="{esc(t_ext)}" '
                        f'data-room="{esc(room_txt)}" '
                        f'data-time="{esc(time_text)}" '
                        f'data-week="{esc(week)}" '
                        f'data-code="{esc(getattr(c, "course_code", ""))}" '
                        f'data-summary="{esc(getattr(c, "course_summary_ch", ""))}" '
                        f'style="cursor:pointer;">{esc(getattr(c, "course_name", ""))}</div>'
                        f'<div class="course-room">{esc(dept_display(getattr(c, "department_code", "")))}</div>'
                        f'<div class="course-room">{esc(getattr(c, "teacher", ""))} {esc(room_txt)}</div>'
                    )
                )

            table_html += "<td>" + "<hr style='border:none;border-top:1px solid #e5e7eb;margin:8px 0;'>".join(parts) + "</td>"

        table_html += "</tr>"

    table_html += "</table></div>"
    return table_html


def add_course(request):
    fixed_semester = "1141"

    role_name, _ = get_user_display_name(request.user)
    admin_mode = request.user.is_authenticated and role_name == "老師"
    if not admin_mode:
        return redirect("course_query")

    if request.method == "POST":
        form = CourseForm(request.POST)
        if form.is_valid():
            c = form.save(commit=False)
            c.semester = fixed_semester
            c.teacher = "連中岳"
            c.save()
            return redirect(f"{reverse('course_query')}?semester={fixed_semester}&submitted=1")
    else:
        form = CourseForm()

    return render(request, "main/add_course.html", {"form": form})


@require_POST
def delete_course(request, course_id: int):
    role_name, _ = get_user_display_name(request.user)
    admin_mode = request.user.is_authenticated and role_name == "老師"
    if not admin_mode:
        return redirect("course_query")

    Course.objects.filter(
        id=course_id,
        semester="1141",
        teacher__icontains="連中岳",
    ).delete()

    return redirect(f"{reverse('course_query')}?semester=1141&submitted=1")


# ==============================
#        課程查詢 + 顯示
# ==============================


def course_query(request):
    ensure_default_accounts()
    login_error = ""
    conflicts = []
    # ✅ 0) 先處理「學生 AJAX：新增/移除個人課表」避免誤進登入判斷
    if request.method == "POST" and safe_str(request.POST.get("action")) in {"add_my_course", "remove_my_course"}:
        if not request.user.is_authenticated or not Student.objects.filter(user=request.user).exists():
            return JsonResponse({"ok": False, "message": "請先以學生身分登入。"}, status=401)

        action = safe_str(request.POST.get("action"))
        course_id_raw = safe_str(request.POST.get("course_id"))
        force = safe_str(request.POST.get("force")) == "1"

        try:
            course_id = int(course_id_raw)
        except Exception:
            return JsonResponse({"ok": False, "message": "course_id 格式錯誤。"}, status=400)

        c = Course.objects.filter(id=course_id).first()
        if not c:
            return JsonResponse({"ok": False, "message": "找不到課程。"}, status=404)

        ensure_fixed_personal_courses(request)
        ids = _get_personal_ids(request)
        id_set = set(ids)
        existing_courses = list(Course.objects.filter(id__in=id_set))
        conflicts = _conflict_slots(existing_courses, c)
        if action == "remove_my_course":
            if is_required_course_id(course_id):
                return JsonResponse(
                    {"ok": False, "required": True, "message": required_remove_message(course_id)},
                    status=409,
                )

            if course_id in id_set:
                ids = [x for x in ids if x != course_id]
                _set_personal_ids(request, ids)
            return JsonResponse({"ok": True, "message": "已從個人課表移除。"})

        if course_id in id_set:
            return JsonResponse({"ok": True, "message": "此課程已在個人課表中。"})

        existing_courses = list(Course.objects.filter(id__in=list(id_set)))
        if conflicts:
            return JsonResponse(
                {
                    "ok": False,
                    "conflict": True,
                    "": conflicts,
                    "message": f"此課程與你的個人課表衝堂：{_format_conflicts(conflicts)}",
                },
                status=409,
            )

        ids.append(course_id)
        _set_personal_ids(request, ids)
        return JsonResponse({"ok": True, "message": "已新增到個人課表。", "warning": bool(conflicts), "conflicts": conflicts})

    # ✅ 1) 再處理「登入（POST）」：必須 username/password 都非空才算登入
    if request.method == "POST":
        username_in = safe_str(request.POST.get("username"))
        password_in = request.POST.get("password") or ""
        role_in = safe_str(request.POST.get("role"))  # admin / student

        if username_in and password_in and role_in in {"admin", "student"}:
            user = authenticate(request, username=username_in, password=password_in)
            if user is None:
                login_error = "帳號或密碼錯誤"
            else:
                ok = True
                if role_in == "student":
                    if Student.objects.filter(user=user).first() is None:
                        ok = False
                        login_error = "此帳號不是學生身分"
                elif role_in == "admin":
                    if not (user.is_staff or user.is_superuser or Teacher.objects.filter(user=user).exists()):
                        ok = False
                        login_error = "此帳號不是管理員/老師身分"

                if ok:
                    login(request, user)
                    return redirect("course_query")

    ensure_courses_loaded()

    def norm(v):
        return safe_str(v)

    period_time_map = {
        1: "08:10~09:00",
        2: "09:10~10:00",
        3: "10:10~11:00",
        4: "11:10~12:00",
        5: "12:40~13:30",
        6: "13:40~14:30",
        7: "14:40~15:30",
        8: "15:40~16:30",
        9: "16:40~17:30",
        10: "17:40~18:30",
        11: "18:35~19:25",
        12: "19:30~20:20",
        13: "20:25~21:15",
        14: "21:20~22:10",
    }

    day_labels = [("1", "一"), ("2", "二"), ("3", "三"), ("4", "四"), ("5", "五"), ("6", "六"), ("7", "日")]

    class_type_options = (
        Course.objects.exclude(division__isnull=True)
        .exclude(division__exact="")
        .values_list("division", flat=True)
        .distinct()
        .order_by("division")
    )

    role_name, display_name = get_user_display_name(request.user)
    admin_mode = request.user.is_authenticated and role_name == "老師"

    submitted = request.GET.get("submitted")

    semester = norm(request.GET.get("semester"))
    system = norm(request.GET.get("system"))
    grade = norm(request.GET.get("grade"))
    department = norm(request.GET.get("department"))
    teacher = norm(request.GET.get("teacher"))
    course_name = norm(request.GET.get("course_name"))
    course_code = norm(request.GET.get("course_code"))
    class_type = norm(request.GET.get("class_type"))

    days_selected = request.GET.getlist("day")
    periods_selected = request.GET.getlist("period")

    period_range = list(range(1, 15))
    periods_selected_int = []
    for p in periods_selected:
        try:
            periods_selected_int.append(int(p))
        except Exception:
            continue

    total_count = Course.objects.count()

    personal_timetable_html = ""
    my_course_ids = []
    if request.user.is_authenticated and Student.objects.filter(user=request.user).exists():
        ensure_fixed_personal_courses(request)
        my_course_ids = _get_personal_ids(request)
        if my_course_ids:
            m = {c.id: c for c in Course.objects.filter(id__in=my_course_ids)}
            personal_courses = [m[i] for i in my_course_ids if i in m]
            personal_timetable_html = build_grid_timetable_html(personal_courses, title="我的個人課表")
        else:
            personal_timetable_html = '<div class="no-result">找不到 A0 的「系統分析 / 研究概論(資管系)」課程資料。</div>'

    timetable_html = ""

    # ==============================
    # ✅ 管理員模式
    # ==============================
    if admin_mode:
        admin_courses = []
        if submitted and semester:
            qs = Course.objects.filter(teacher__icontains="連中岳", semester=semester)

            def course_sort_key(c):
                return (norm(c.day) or "9", norm(c.period) or "", norm(c.course_name) or "")

            admin_courses = sorted(list(qs), key=course_sort_key)
            for c in admin_courses:
                c.dept_name = dept_display(getattr(c, "department_code", ""))

        context = {
            "semester": semester,
            "system": system,
            "grade": grade,
            "department": department,
            "teacher": teacher,
            "course_name": course_name,
            "course_code": course_code,
            "class_type": class_type,
            "class_type_options": list(class_type_options),
            "days_selected": days_selected,
            "period_range": period_range,
            "periods_selected_int": periods_selected_int,
            "total_count": total_count,
            "timetable_html": "",
            "role_name": role_name,
            "display_name": display_name,
            "login_error": login_error,
            "personal_timetable_html": personal_timetable_html,
            "admin_mode": True,
            "admin_courses": admin_courses,
            "my_course_ids": my_course_ids,
            "current_full_path": request.get_full_path(),
            "building_url_map": json.dumps(BUILDING_URL_MAP),
        }
        return render(request, "main/course_query.html", context)

    # ==============================
    # ✅ 學生模式
    # ==============================
    courses = []

    only_semester = bool(semester) and not any(
        [system, grade, department, teacher, course_name, course_code, class_type, days_selected, periods_selected]
    )

    no_condition = not any(
        [semester, system, grade, department, teacher, course_name, course_code, class_type, days_selected, periods_selected]
    )

    if not submitted:
        timetable_html = '<div class="no-result">尚未查詢，請先設定條件後按「查詢」。</div>'
    else:
        if no_condition:
            timetable_html = (
                '<div class="no-result" style="color:#b91c1c;">'
                "請至少選擇一個查詢條件（例如學期、科系、老師、星期或節次）再按「查詢」。"
                "</div>"
            )
        else:
            qs = Course.objects.all()

            if semester:
                qs = qs.filter(semester=semester)
            if system:
                qs = apply_system_filter(qs, system)
            if department:
                qs = qs.filter(department_code__exact=department)
            if grade:
                qs = qs.filter(grade=grade)
            if teacher:
                qs = qs.filter(teacher__icontains=teacher)
            if course_name:
                qs = qs.filter(course_name__icontains=course_name)
            if course_code:
                qs = qs.filter(course_code__icontains=course_code)
            if class_type:
                qs = qs.filter(division__icontains=class_type)
            if days_selected:
                qs = qs.filter(day__in=days_selected)

            courses = list(qs)

            if periods_selected:
                try:
                    period_need = {int(p) for p in periods_selected}
                except Exception:
                    period_need = set()

                if period_need:
                    filtered_list = []
                    for c in courses:
                        c_periods = set(parse_periods(norm(getattr(c, "period", ""))))
                        if c_periods & period_need:
                            filtered_list.append(c)
                    courses = filtered_list

            timetable = {}
            for c in courses:
                day_str = norm(c.day)
                periods_raw = norm(c.period)
                if not day_str or not periods_raw:
                    continue

                for pp in parse_periods(periods_raw):
                    timetable.setdefault(day_str, {}).setdefault(pp, []).append(c)

            if only_semester:
                if courses:

                    def course_sort_key(c):
                        return (norm(c.day) or "9", norm(c.period), norm(c.department_code), norm(c.class_group))

                    courses_sorted = sorted(courses, key=course_sort_key)
                    day_map = {"1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "日"}

                    list_html = '<div class="timetable-wrapper">'
                    list_html += f'<div class="timetable-title">學期 {esc(semester)} 課程列表</div>'
                    list_html += '<table class="timetable">'
                    list_html += (
                        "<tr>"
                        "<th>課別名稱</th>"
                        "<th>系所</th>"
                        "<th>科目名稱</th>"
                        "<th>老師</th>"
                        "<th>教室地點</th>"
                        "<th>時間</th>"
                        "</tr>"
                    )

                    for c in courses_sorted:
                        dept_name = dept_display(c.department_code)
                        day_ch = day_map.get(norm(c.day), norm(c.day) or "-")
                        period_str = norm(c.period) or "-"
                        week_info = norm(c.week_info)

                        time_text = f"星期{day_ch} 第{period_str}節"
                        if week_info:
                            time_text += f"（{week_info}）"

                        t_ch, t_cat, t_ext = teacher_meta_for_course(c)
                        room_txt = room_display(c)

                        name_html = (
                            f'<span class="course-clickable" '
                            f'data-id="{c.id}" '
                            f'data-name="{esc(c.course_name)}" '
                            f'data-dept="{esc(dept_name)}" '
                            f'data-teacher="{esc(c.teacher)}" '
                            f'data-teacher-ch="{esc(t_ch)}" '
                            f'data-teacher-category="{esc(t_cat)}" '
                            f'data-teacher-ext="{esc(t_ext)}" '
                            f'data-room="{esc(room_txt)}" '
                            f'data-time="{esc(time_text)}" '
                            f'data-week="{esc(c.week_info)}" '
                            f'data-code="{esc(c.course_code)}" '
                            f'data-summary="{esc(c.course_summary_ch)}" '
                            f'style="cursor:pointer;">{esc(c.course_name)}</span>'
                        )

                        list_html += (
                            "<tr>"
                            f"<td>{esc(c.division) or '-'}</td>"
                            f"<td>{esc(dept_name)}</td>"
                            f"<td>{name_html}</td>"
                            f"<td>{esc(c.teacher) or '-'}</td>"
                            f"<td>{esc(room_txt)}</td>"
                            f"<td>{esc(time_text)}</td>"
                            "</tr>"
                        )

                    list_html += "</table></div>"
                    timetable_html = list_html
                else:
                    timetable_html = '<div class="no-result">此學期目前查無任何課程資料。</div>'

            else:
                if courses:
                    table_html = '<div class="timetable-wrapper">'
                    table_html += '<div class="timetable-title">課表</div>'
                    table_html += '<table class="timetable">'
                    table_html += "<tr><th>節次</th>"
                    for _val, label in day_labels:
                        table_html += f"<th>星期{esc(label)}</th>"
                    table_html += "</tr>"

                    for p in period_range:
                        t = period_time_map.get(p, "")
                        th_html = f'{p}<div style="font-size:11px;color:#6b7280;margin-top:4px;">{esc(t)}</div>' if t else f"{p}"
                        table_html += f"<tr><th>{th_html}</th>"

                        for day_val, day_label in day_labels:
                            if days_selected and day_val not in days_selected:
                                table_html += "<td>&nbsp;</td>"
                                continue

                            courses_in_cell = timetable.get(day_val, {}).get(p, [])
                            if not courses_in_cell:
                                table_html += "<td>&nbsp;</td>"
                                continue

                            if len(courses_in_cell) <= 2:
                                parts = []
                                for c in courses_in_cell:
                                    time_text = f"星期{day_label} 第{esc(c.period) or '-'}節"
                                    if norm(c.week_info):
                                        time_text += f"（{esc(c.week_info)}）"

                                    t_ch, t_cat, t_ext = teacher_meta_for_course(c)
                                    room_txt = room_display(c)

                                    parts.append(
                                        (
                                            f'<div class="course-cell course-clickable" '
                                            f'data-id="{c.id}" '
                                            f'data-name="{esc(c.course_name)}" '
                                            f'data-dept="{esc(dept_display(c.department_code))}" '
                                            f'data-teacher="{esc(c.teacher)}" '
                                            f'data-teacher-ch="{esc(t_ch)}" '
                                            f'data-teacher-category="{esc(t_cat)}" '
                                            f'data-teacher-ext="{esc(t_ext)}" '
                                            f'data-room="{esc(room_txt)}" '
                                            f'data-time="{esc(time_text)}" '
                                            f'data-week="{esc(c.week_info)}" '
                                            f'data-code="{esc(c.course_code)}" '
                                            f'data-summary="{esc(c.course_summary_ch)}" '
                                            f'style="cursor:pointer;">{esc(c.course_name)}</div>'
                                            f'<div class="course-room">{esc(dept_display(c.department_code))}</div>'
                                            f'<div class="course-room">{esc(c.teacher)} {esc(room_txt)}</div>'
                                        )
                                    )

                                table_html += "<td>" + "<hr style='border:none;border-top:1px solid #e5e7eb;margin:8px 0;'>".join(parts) + "</td>"
                                continue

                            cell_id = f"cell_{day_val}_{p}"
                            first = courses_in_cell[0]

                            select_html = (
                                f'<select class="cell-select" id="{cell_id}_select" '
                                f"onchange=\"updateTimetableCell('{cell_id}');\" "
                                f'title="{esc(first.course_name)}">'
                            )

                            for idx, c in enumerate(courses_in_cell):
                                cname = norm(c.course_name)
                                opt_label = cname if len(cname) <= 18 else cname[:18] + "…"
                                selected = "selected" if idx == 0 else ""

                                t_ch, t_cat, t_ext = teacher_meta_for_course(c)
                                room_txt = room_display(c)

                                time_text = f"星期{day_label} 第{norm(getattr(c, 'period', '')) or '-'}節"
                                if norm(c.week_info):
                                    time_text += f"（{norm(c.week_info)}）"

                                select_html += (
                                    f'<option value="{idx}" {selected} title="{esc(cname)}" '
                                    f'data-id="{c.id}" '
                                    f'data-name="{esc(c.course_name)}" '
                                    f'data-dept="{esc(dept_display(c.department_code))}" '
                                    f'data-teacher="{esc(c.teacher)}" '
                                    f'data-teacher-ch="{esc(t_ch)}" '
                                    f'data-teacher-category="{esc(t_cat)}" '
                                    f'data-teacher-ext="{esc(t_ext)}" '
                                    f'data-room="{esc(room_txt)}" '
                                    f'data-time="{esc(time_text)}" '
                                    f'data-week="{esc(c.week_info)}" '
                                    f'data-code="{esc(c.course_code)}" '
                                    f'data-summary="{esc(c.course_summary_ch)}" '
                                    f'>{esc(opt_label)}</option>'
                                )

                            select_html += "</select>"

                            display_html = (
                                f'<div class="cell-display" id="{cell_id}_display">'
                                f'<div class="course-cell">{esc(first.course_name)}</div>'
                                f'<div class="course-room">{esc(dept_display(first.department_code))}</div>'
                                f'<div class="course-room">{esc(first.teacher)} {esc(room_display(first))}</div>'
                                f"</div>"
                            )
                            table_html += f"<td>{select_html}{display_html}</td>"

                        table_html += "</tr>"

                    table_html += "</table></div>"
                    timetable_html = table_html
                else:
                    timetable_html = '<div class="no-result">查無符合條件的課程，請調整查詢條件再試一次。</div>'

    context = {
        "semester": semester,
        "system": system,
        "grade": grade,
        "department": department,
        "teacher": teacher,
        "course_name": course_name,
        "course_code": course_code,
        "class_type": class_type,
        "class_type_options": list(class_type_options),
        "days_selected": days_selected,
        "period_range": period_range,
        "periods_selected_int": periods_selected_int,
        "total_count": total_count,
        "timetable_html": timetable_html,
        "role_name": role_name,
        "display_name": display_name,
        "login_error": login_error,
        "personal_timetable_html": personal_timetable_html,
        "admin_mode": False,
        "my_course_ids": my_course_ids,
        "current_full_path": request.get_full_path(),
        "building_url_map": json.dumps(BUILDING_URL_MAP),
    }
    return render(request, "main/course_query.html", context)


# ==============================
# ✅ Excel 匯入工具
# ==============================


def import_excel(request):
    file_path = EXCEL_DIR / "課程查詢_1131.xlsx"
    try:
        df = pd.read_excel(file_path, header=4)
    except Exception as e:
        return HttpResponse(f"讀取 Excel 檔案失敗：{e}")

    count = _import_df_to_course(df)
    return HttpResponse(f"匯入完成（單一檔案），共匯入 {count} 筆資料！")


def import_all_excels(request):
    excel_files = sorted(EXCEL_DIR.glob("*.xlsx"))
    if not excel_files:
        return HttpResponse(f"在 {EXCEL_DIR} 沒有找到任何 .xlsx 檔案，請確認路徑。")

    total_files = 0
    total_rows = 0
    log_messages = []

    for file_path in excel_files:
        try:
            df = pd.read_excel(file_path, header=4)
            count = _import_df_to_course(df)
            total_files += 1
            total_rows += count
            log_messages.append(f"{file_path.name}：{count} 筆")
        except Exception as e:
            log_messages.append(f"{file_path.name} 讀取失敗：{e}")

    detail = "<br>".join(log_messages)
    return HttpResponse(
        f"匯入完成，共處理 {total_files} 個檔案，總共 {total_rows} 筆資料。<br><br>{detail}"
    )


# ==============================
# ✅ teacher_info + backfill（保留）
# ==============================


@require_GET
def teacher_info(request):
    name = safe_str(request.GET.get("name"))
    if not name:
        return JsonResponse({"ok": False, "message": "缺少 name"}, status=400)

    t = Teacher.objects.filter(name_ch=name).first()
    if not t:
        t = Teacher.objects.filter(name_en=name).first()

    if not t:
        return JsonResponse({"ok": False, "message": "找不到教師資料"}, status=404)

    name_ch = safe_str(getattr(t, "name_ch", "")) or name

    category = (
        safe_str(getattr(t, "category", "")) or
        safe_str(getattr(t, "type", "")) or
        safe_str(getattr(t, "title", "")) or
        safe_str(getattr(t, "role", ""))
    )

    ext = (
        safe_str(getattr(t, "extension", "")) or
        safe_str(getattr(t, "ext", "")) or
        safe_str(getattr(t, "phone_ext", "")) or
        safe_str(getattr(t, "office_ext", ""))
    )

    return JsonResponse({
        "ok": True,
        "name_ch": name_ch or "-",
        "category": category or "-",
        "ext": ext or "-",
    })


@require_GET
def backfill_classroom_from_excel(request):
    excel_files = sorted(EXCEL_DIR.glob("*.xlsx"))
    if not excel_files:
        return HttpResponse(f"在 {EXCEL_DIR} 沒有找到任何 .xlsx 檔案。")

    idx = {}
    loaded_rows = 0

    for file_path in excel_files:
        try:
            df = pd.read_excel(file_path, header=4)
        except Exception:
            continue

        for _, row in df.iterrows():
            sem = safe_get(row, "學期")
            code = safe_get(row, "科目代碼(新碼全碼)")
            name = safe_get(row, "科目中文名稱")
            room = room_from_row(row)
            if sem and (code or name) and room:
                idx[(sem, code, name)] = room
                loaded_rows += 1

    qs = Course.objects.filter(Q(classroom__isnull=True) | Q(classroom__exact=""))
    updated = 0

    for c in qs:
        sem = safe_str(getattr(c, "semester", ""))
        code = safe_str(getattr(c, "course_code", ""))
        name = safe_str(getattr(c, "course_name", ""))
        room = idx.get((sem, code, name))
        if room:
            c.classroom = room
            c.save(update_fields=["classroom"])
            updated += 1

    return HttpResponse(f"回填完成：Excel索引 {loaded_rows} 筆；更新 classroom {updated} 筆。")
# ==============================
# ✅ 舊網址相容：personal/ & personal/remove（給 urls.py 用）
# ==============================

def _course_conflicts(new_course: Course, personal_courses):
    new_day = safe_str(getattr(new_course, "day", ""))
    new_periods = set(parse_periods(safe_str(getattr(new_course, "period", ""))))
    if not new_day or not new_periods:
        return []

    conflicts = []
    for c in personal_courses:
        c_day = safe_str(getattr(c, "day", ""))
        c_periods = set(parse_periods(safe_str(getattr(c, "period", ""))))
        if not c_day or not c_periods:
            continue
        if c_day == new_day and (new_periods & c_periods):
            conflicts.append(c)
    return conflicts


@require_POST
def add_personal_course(request, course_id: int):
    # 只允許學生
    if not request.user.is_authenticated or not Student.objects.filter(user=request.user).exists():
        return JsonResponse({"ok": False, "message": "只有學生可以新增個人課表。"}, status=403)
    try:
      cid = int(course_id)
    except Exception:
      return JsonResponse({"ok": False, "message": "course_id 格式錯誤。"}, status=400)
    c = Course.objects.filter(id=course_id).first()
    if not c:
        return JsonResponse({"ok": False, "message": "找不到課程。"}, status=404)

    ensure_fixed_personal_courses(request)
    ids = _get_personal_ids(request)

    if c.id in ids:
        return JsonResponse({"ok": True, "message": "課程已在個人課表中。", "my_course_ids": ids})

    personal_courses = list(Course.objects.filter(id__in=ids))
    conflicts = _course_conflicts(c, personal_courses)

    force = safe_str(request.POST.get("force"))
    if conflicts and force != "1":
        # 回傳衝堂細節（維持你原本格式）
        conflict_list = []
        day_map = {"1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "日"}
        for cc in conflicts:
            conflict_list.append(
                {
                    "id": cc.id,
                    "name": safe_str(getattr(cc, "course_name", "")),
                    "day": day_map.get(safe_str(getattr(cc, "day", "")), safe_str(getattr(cc, "day", ""))),
                    "period": safe_str(getattr(cc, "period", "")),
                    "teacher": safe_str(getattr(cc, "teacher", "")),
                }
            )

        return JsonResponse(
            {
                "ok": False,
                "conflict": True,
                "message": "新增課程與個人課表衝堂，是否仍要新增？",
                "conflicts": conflict_list,
                "course_id": c.id,
            },
            status=409,
        )

    ids.append(c.id)
    _set_personal_ids(request, ids)

    return JsonResponse(
        {"ok": True, "message": "已新增到個人課表。", "my_course_ids": ids, "warning": bool(conflicts)}
    )


@require_POST
def remove_personal_course(request, course_id: int):
    # 只允許學生
    if not request.user.is_authenticated or not Student.objects.filter(user=request.user).exists():
        return JsonResponse({"ok": False, "message": "只有學生可以移除個人課表。"}, status=403)

    try:
        cid = int(course_id)
    except Exception:
        return JsonResponse({"ok": False, "message": "course_id 格式錯誤。"}, status=400)

    ensure_fixed_personal_courses(request)

    if is_required_course_id(cid):
        return JsonResponse(
            {"ok": False, "required": True, "message": required_remove_message(cid)},
            status=409,
        )

    ids = _get_personal_ids(request)
    ids = [i for i in ids if i != cid]
    _set_personal_ids(request, ids)

    return JsonResponse({"ok": True, "message": "已從個人課表移除。", "my_course_ids": ids})
