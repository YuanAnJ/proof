from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse, JsonResponse
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.db.models import Q
from django.db import transaction
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
import json
import re
from .models import FormsData  # 确保导入你的数据模型
from datetime import datetime
import pandas as pd
from forms.utils.docx_generate import generate_doc
from forms.utils.docx_import import parse_docx_file

# Create your views here.

STANDARD_YEARS = ['2016','2017','2018','2019','2020','2021','2022','2023','2024','2025','2026','2027','2028','2029','2030','2031','2032','2033','2034','2035']  # 可根据需要调整年份范围
STANDARD_ACCEPT_LEVELS = ['中办', '国办', '中央有关部门', '国家部委', '省级','市厅级'] #可根据需要调整采纳级别列表
STANDARD_INSTRUCTION_LEVELS = ['中央主要领导', '中央政治局常委', '中央政治局委员', '中央有关领导','正部级领导','副部级领导','省级有关领导','市级主要领导','市有关领导'] #可根据需要调整批示级别列表

def index(request):
    data = FormsData.objects.all()
    number_year=FormsData.get_number_year(data[0])
    number_seq = FormsData.get_number_seq(data[0])
    return HttpResponse("Hello, world. You're at the forms index.\n{}\n{}".format(number_year,number_seq))

# 创建基本表单视图
def add_form_template(request):
    return render(request, 'forms/form.html')


def batch_import_template(request):
    return render(request, 'forms/batch_import.html', {
        'standard_accept_levels': STANDARD_ACCEPT_LEVELS,
        'standard_instruction_levels': STANDARD_INSTRUCTION_LEVELS,
    })


def _split_values(value):
    if value is None:
        return []
    return [item.strip() for item in re.split(r'[、,，;；]+', str(value)) if item.strip()]


def _normalize_levels(raw_value, standard_levels):
    values = _split_values(raw_value)
    primary = []
    other = []

    for item in values:
        if item in standard_levels or item == '无':
            if item not in primary:
                primary.append(item)
        else:
            if item not in other:
                other.append(item)

    return '、'.join(primary), '、'.join(other)


def _validate_feedback_date(value):
    date_str = str(value or '').strip()
    if not date_str:
        return False, '反馈日期不能为空'

    if re.match(r'^\d{4}-\d{1,2}$', date_str):
        return False, '反馈日期仅到月份，请补全到具体日期'

    try:
        datetime.strptime(date_str, '%Y-%m-%d')
        return True, ''
    except ValueError:
        return False, '反馈日期格式错误，需为YYYY-MM-DD'


def _build_duplicate_key(row):
    return '||'.join([
        str(row.get('title', '')).strip(),
        str(row.get('feedback_date', '')).strip(),
        str(row.get('author', '')).strip(),
        str(row.get('accept_level', '')).strip(),
        str(row.get('instruction_level', '')).strip(),
    ])


def _validate_row_data(row):
    errors = []

    required_fields = {
        'feedback_department': '反馈部门',
        'title': '报告题目',
        'author': '作者一',
        'unit': '作者一单位',
    }

    for field, display_name in required_fields.items():
        if not str(row.get(field, '')).strip():
            errors.append(f'{display_name}不能为空')

    ok, message = _validate_feedback_date(row.get('feedback_date', ''))
    if not ok:
        errors.append(message)

    for i in range(2, 11):
        author_val = str(row.get(f'author_{i}', '')).strip()
        unit_val = str(row.get(f'unit_{i}', '')).strip()
        if (author_val and not unit_val) or (unit_val and not author_val):
            errors.append(f'作者{i}与单位{i}需同时填写或同时留空')

    return errors


def _generate_number():
    latest_record = FormsData.objects.all().order_by('-id').first()
    year_part = datetime.now().strftime('%y')

    if latest_record:
        if year_part == FormsData.get_number_year(latest_record):
            seq_number = int(FormsData.get_number_seq(latest_record)) + 1
            return f'JCZX{year_part}{seq_number:04d}'
        return f'JCZX{year_part}0001'
    return f'JCZX{year_part}0001'

# TODO：完善搜索逻辑，分页逻辑，跳转逻辑
def query_form_template(request):
    # 获取搜索参数
    search_type = request.GET.get('search_type', 'none')  # 默认为'none'
    search_query = request.GET.get('search_query', '')
    search_accept_level = request.GET.get('search_accept_level', '')
    search_instruction_level = request.GET.get('search_instruction_level', '')
    search_years = request.GET.getlist('search_years')
    page = request.GET.get('page')

    # 构建基础查询集
    data_list = FormsData.objects.filter(is_delete=False)
    
    # 根据搜索类型和关键词筛选
    if search_query and search_type != 'none':
        if search_type == 'unit':
            data_list = data_list.filter(
                Q(unit__icontains=search_query) |
                Q(unit_2__icontains=search_query) |
                Q(unit_3__icontains=search_query) |
                Q(unit_4__icontains=search_query) |
                Q(unit_5__icontains=search_query) |
                Q(unit_6__icontains=search_query) |
                Q(unit_7__icontains=search_query) |
                Q(unit_8__icontains=search_query) |
                Q(unit_9__icontains=search_query) |
                Q(unit_10__icontains=search_query)
            )
        elif search_type == 'author':
            data_list = data_list.filter(
                Q(author__icontains=search_query) |
                Q(author_2__icontains=search_query) |
                Q(author_3__icontains=search_query) |
                Q(author_4__icontains=search_query) |
                Q(author_5__icontains=search_query) |
                Q(author_6__icontains=search_query) |
                Q(author_7__icontains=search_query) |
                Q(author_8__icontains=search_query) |
                Q(author_9__icontains=search_query) |
                Q(author_10__icontains=search_query)
            )
        elif search_type == 'title':
            data_list = data_list.filter(title__icontains=search_query)
        elif search_type == 'feedback_department':
            data_list = data_list.filter(feedback_department__icontains=search_query)
        elif search_type == 'number':
            # 按编号模糊匹配（也可改为精确匹配 number == search_query）
            data_list = data_list.filter(number__icontains=search_query)
    
    # 根据年份筛选
    if search_years:
        data_list = data_list.filter(feedback_date__year__in=map(int, search_years))

    # 根据采纳级别筛选
    if search_accept_level:
        accept_levels = search_accept_level.split('、')
        accept_query = Q()
        for level in accept_levels:
            accept_query |= Q(accept_level__icontains=level)
        data_list = data_list.filter(accept_query)
    
    # 根据批示级别筛选
    if search_instruction_level:
        instruction_levels = search_instruction_level.split('、')
        instruction_query = Q()
        for level in instruction_levels:
            instruction_query |= Q(instruction_level__icontains=level)
        data_list = data_list.filter(instruction_query)
    
    # 设置每页显示10条数据
    paginator = Paginator(data_list, 10)
    
    try:
        data = paginator.page(page)
    except PageNotAnInteger:
        data = paginator.page(1)
    except EmptyPage:
        data = paginator.page(paginator.num_pages)
    
    return render(request, 'forms/query.html', {
        'data': data,
        'search_type': search_type,
        'search_query': search_query,
        'search_accept_level': search_accept_level,
        'search_instruction_level': search_instruction_level,
        'search_years':search_years,
        'standard_years': STANDARD_YEARS,
        'standard_accept_levels':STANDARD_ACCEPT_LEVELS,
        'standard_instruction_levels':STANDARD_INSTRUCTION_LEVELS
    })

def edit_form_template(request, number):
    # 获取要编辑的数据记录
    data = get_object_or_404(FormsData, number=number)
    
    # GET 请求显示编辑表单
    return render(request, 'forms/edit.html', {
        'data': data
    })

def generate_doc_api(request,number):
    # 获取要生成文档的数据记录
    data = get_object_or_404(FormsData, number=number)

    # print(type(data.instruction_level))

    file_name = generate_doc(data)

    response = HttpResponse(open('media/proof_file.docx', 'rb'), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = f'attachment; filename="{file_name}_Proof.docx"'
    return response

def export_form_template(request):
    """导出信息页面视图"""
    # 获取筛选参数
    feedback_date_start = request.GET.get('feedback_date_start', '')
    feedback_date_end = request.GET.get('feedback_date_end', '')
    search_feedback_department = request.GET.get('search_feedback_department', '')
    search_authors = request.GET.get('search_authors', '')
    search_accept_level = request.GET.get('search_accept_level', '')
    search_instruction_level = request.GET.get('search_instruction_level', '')
    page = request.GET.get('page')

    # 构建基础查询集
    data_list = FormsData.objects.filter(is_delete=False)
    
    # 按反馈日期范围筛选
    if feedback_date_start:
        data_list = data_list.filter(feedback_date__gte=feedback_date_start)
    if feedback_date_end:
        data_list = data_list.filter(feedback_date__lte=feedback_date_end)
    
    # 按反馈部门筛选
    if search_feedback_department:
        departments = [dept.strip() for dept in search_feedback_department.split(',')]
        dept_query = Q()
        for dept in departments:
            dept_query |= Q(feedback_department__icontains=dept)
        data_list = data_list.filter(dept_query)
    
    # 按作者筛选
    if search_authors:
        authors = [author.strip() for author in search_authors.split(',')]
        author_query = Q()
        for author in authors:
            author_query |= Q(author__icontains=author) | \
                           Q(author_2__icontains=author) | \
                           Q(author_3__icontains=author) | \
                           Q(author_4__icontains=author) | \
                           Q(author_5__icontains=author) | \
                           Q(author_6__icontains=author) | \
                           Q(author_7__icontains=author) | \
                           Q(author_8__icontains=author) | \
                           Q(author_9__icontains=author) | \
                           Q(author_10__icontains=author)
        data_list = data_list.filter(author_query)
    
    # 按采纳级别筛选
    if search_accept_level:
        accept_levels = [level.strip() for level in search_accept_level.split(',')]
        accept_query = Q()
        for level in accept_levels:
            accept_query |= Q(accept_level__icontains=level)
        data_list = data_list.filter(accept_query)
    
    # 按批示级别筛选
    if search_instruction_level:
        instruction_levels = [level.strip() for level in search_instruction_level.split(',')]
        instruction_query = Q()
        for level in instruction_levels:
            instruction_query |= Q(instruction_level__icontains=level)
        data_list = data_list.filter(instruction_query)
    
    # 设置每页显示10条数据
    paginator = Paginator(data_list, 10)
    
    try:
        data = paginator.page(page)
    except PageNotAnInteger:
        data = paginator.page(1)
    except EmptyPage:
        data = paginator.page(paginator.num_pages)
    
    return render(request, 'forms/export.html', {
        'data': data,
        'feedback_date_start': feedback_date_start,
        'feedback_date_end': feedback_date_end,
        'search_feedback_department': search_feedback_department,
        'search_authors': search_authors,
        'search_accept_level': search_accept_level,
        'search_instruction_level': search_instruction_level,
        'standard_accept_levels': STANDARD_ACCEPT_LEVELS,
        'standard_instruction_levels': STANDARD_INSTRUCTION_LEVELS,
    })

def export_excel_api(request):
    # 获取筛选参数
    feedback_date_start = request.GET.get('feedback_date_start', '')
    feedback_date_end = request.GET.get('feedback_date_end', '')
    search_feedback_department = request.GET.get('search_feedback_department', '')
    search_authors = request.GET.get('search_authors', '')
    search_accept_level = request.GET.get('search_accept_level', '')
    search_instruction_level = request.GET.get('search_instruction_level', '')
    
    # 构建基础查询集
    data = FormsData.objects.filter(is_delete=False)
    
    # 应用与export_form相同的筛选逻辑
    if feedback_date_start:
        data = data.filter(feedback_date__gte=feedback_date_start)
    if feedback_date_end:
        data = data.filter(feedback_date__lte=feedback_date_end)
    
    if search_feedback_department:
        departments = [dept.strip() for dept in search_feedback_department.split(',')]
        dept_query = Q()
        for dept in departments:
            dept_query |= Q(feedback_department__icontains=dept)
        data = data.filter(dept_query)
    
    if search_authors:
        authors = [author.strip() for author in search_authors.split(',')]
        author_query = Q()
        for author in authors:
            author_query |= Q(author__icontains=author) | \
                           Q(author_2__icontains=author) | \
                           Q(author_3__icontains=author) | \
                           Q(author_4__icontains=author) | \
                           Q(author_5__icontains=author) | \
                           Q(author_6__icontains=author) | \
                           Q(author_7__icontains=author) | \
                           Q(author_8__icontains=author) | \
                           Q(author_9__icontains=author) | \
                           Q(author_10__icontains=author)
        data = data.filter(author_query)
    
    if search_accept_level:
        accept_levels = [level.strip() for level in search_accept_level.split(',')]
        accept_query = Q()
        for level in accept_levels:
            accept_query |= Q(accept_level__icontains=level)
        data = data.filter(accept_query)
    
    if search_instruction_level:
        instruction_levels = [level.strip() for level in search_instruction_level.split(',')]
        instruction_query = Q()
        for level in instruction_levels:
            instruction_query |= Q(instruction_level__icontains=level)
        data = data.filter(instruction_query)
    
    # 创建数据列表
    data_list = []
    for item in data:
        data_list.append({
            '编号': item.number,
            '反馈日期': item.feedback_date.strftime('%Y-%m-%d'),
            '反馈部门': item.feedback_department,
            '标题': item.title,
            '作者一': item.author,
            '作者一单位': item.unit,
            '作者二': item.author_2 or '',
            '作者二单位': item.unit_2 or '',
            '作者三': item.author_3 or '',
            '作者三单位': item.unit_3 or '',
            '作者四': item.author_4 or '',
            '作者四单位': item.unit_4 or '',
            '作者五': item.author_5 or '',
            '作者五单位': item.unit_5 or '',
            '作者六': item.author_6 or '',
            '作者六单位': item.unit_6 or '',
            '作者七': item.author_7 or '',
            '作者七单位':item.unit_7 or '',
            '作者八': item.author_8 or '',
            '作者八单位': item.unit_8 or '',
            '作者九': item.author_9 or '',
            '作者九单位': item.unit_9 or '',
            '作者十': item.author_10 or '',
            '作者十单位': item.unit_10 or '',
            '采纳级别': item.accept_level,
            '批示级别': item.instruction_level,
            '备注': item.remark or ''
        })
    
    # 创建DataFrame
    df = pd.DataFrame(data_list)
    
    # 生成Excel文件
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'backup_{timestamp}.xlsx'
    
    # 创建HTTP响应
    response = HttpResponse(content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    
    # 将DataFrame写入Excel
    df.to_excel(response, index=False, engine='openpyxl')
    
    return response


@csrf_exempt
@require_http_methods(["POST"])
def batch_import_preview_api(request):
    try:
        files = request.FILES.getlist('files')
        if not files:
            return JsonResponse({
                'status': 'error',
                'message': '请至少上传一个docx文件'
            }, status=400)

        rows = []
        file_errors = []
        for index, uploaded_file in enumerate(files):
            if not uploaded_file.name.lower().endswith('.docx'):
                file_errors.append({
                    'file': uploaded_file.name,
                    'error': '仅支持docx文件'
                })
                continue

            try:
                parsed = parse_docx_file(uploaded_file)
                accept_level, other_accept_level = _normalize_levels(
                    parsed.get('accept_level', ''),
                    STANDARD_ACCEPT_LEVELS,
                )
                instruction_level, other_instruction_level = _normalize_levels(
                    parsed.get('instruction_level', ''),
                    STANDARD_INSTRUCTION_LEVELS,
                )

                row = {
                    'row_id': index,
                    'source_file': uploaded_file.name,
                    'feedback_date': parsed.get('feedback_date', ''),
                    'feedback_date_raw': parsed.get('feedback_date_raw', ''),
                    'is_partial_date': bool(parsed.get('is_partial_date', False)),
                    'feedback_department': parsed.get('feedback_department', ''),
                    'title': parsed.get('title', ''),
                    'author': parsed.get('author', ''),
                    'unit': parsed.get('unit', ''),
                    'accept_level': accept_level,
                    'other_accept_level': other_accept_level,
                    'instruction_level': instruction_level,
                    'other_instruction_level': other_instruction_level,
                    'remark': parsed.get('remark', ''),
                }

                for i in range(2, 11):
                    row[f'author_{i}'] = parsed.get(f'author_{i}', '')
                    row[f'unit_{i}'] = parsed.get(f'unit_{i}', '')

                warnings = []
                if row.get('is_partial_date'):
                    warnings.append('反馈日期仅精确到月，请在表格中补全到日')

                row['errors'] = _validate_row_data(row)
                row['warnings'] = warnings
                rows.append(row)
            except Exception as exc:
                file_errors.append({
                    'file': uploaded_file.name,
                    'error': f'解析失败: {exc}'
                })

        return JsonResponse({
            'status': 'success',
            'rows': rows,
            'file_errors': file_errors,
            'summary': {
                'total_files': len(files),
                'parsed_rows': len(rows),
                'file_error_count': len(file_errors),
            }
        })
    except Exception as exc:
        return JsonResponse({
            'status': 'error',
            'message': str(exc),
        }, status=400)


@csrf_exempt
@require_http_methods(["POST"])
def batch_import_confirm_api(request):
    try:
        payload = json.loads(request.body)
        rows = payload.get('rows', [])
        if not isinstance(rows, list) or not rows:
            return JsonResponse({
                'status': 'error',
                'message': '提交数据为空'
            }, status=400)

        results = []
        created_count = 0
        duplicate_count = 0
        invalid_count = 0
        batch_keys = set()

        with transaction.atomic():
            for index, raw_row in enumerate(rows):
                row = dict(raw_row)

                accept_level, other_accept_level = _normalize_levels(
                    row.get('accept_level', ''),
                    STANDARD_ACCEPT_LEVELS,
                )
                instruction_level, other_instruction_level = _normalize_levels(
                    row.get('instruction_level', ''),
                    STANDARD_INSTRUCTION_LEVELS,
                )

                row['accept_level'] = accept_level
                row['other_accept_level'] = other_accept_level
                row['instruction_level'] = instruction_level
                row['other_instruction_level'] = other_instruction_level

                row_errors = _validate_row_data(row)
                duplicate_key = _build_duplicate_key(row)

                if duplicate_key in batch_keys:
                    results.append({
                        'row_id': row.get('row_id', index),
                        'status': 'skipped_duplicate',
                        'reason': '同一批次内重复记录',
                    })
                    duplicate_count += 1
                    continue

                if row_errors:
                    results.append({
                        'row_id': row.get('row_id', index),
                        'status': 'invalid',
                        'reason': '；'.join(row_errors),
                    })
                    invalid_count += 1
                    continue

                exists = FormsData.objects.filter(
                    is_delete=False,
                    title=str(row.get('title', '')).strip(),
                    feedback_date=str(row.get('feedback_date', '')).strip(),
                    author=str(row.get('author', '')).strip(),
                    accept_level=str(row.get('accept_level', '')).strip(),
                    instruction_level=str(row.get('instruction_level', '')).strip(),
                ).exists()

                if exists:
                    results.append({
                        'row_id': row.get('row_id', index),
                        'status': 'skipped_duplicate',
                        'reason': '与已存在记录重复',
                    })
                    duplicate_count += 1
                    batch_keys.add(duplicate_key)
                    continue

                form_data = FormsData(
                    number=_generate_number(),
                    feedback_date=row.get('feedback_date'),
                    feedback_department=row.get('feedback_department', ''),
                    title=row.get('title', ''),
                    author=row.get('author', ''),
                    unit=row.get('unit', ''),
                    accept_level=row.get('accept_level', ''),
                    other_accept_level=row.get('other_accept_level', ''),
                    instruction_level=row.get('instruction_level', ''),
                    other_instruction_level=row.get('other_instruction_level', ''),
                    remark=row.get('remark', ''),
                )

                for i in range(2, 11):
                    setattr(form_data, f'author_{i}', row.get(f'author_{i}', ''))
                    setattr(form_data, f'unit_{i}', row.get(f'unit_{i}', ''))

                form_data.save()
                batch_keys.add(duplicate_key)
                created_count += 1
                results.append({
                    'row_id': row.get('row_id', index),
                    'status': 'created',
                    'number': form_data.number,
                })

        return JsonResponse({
            'status': 'success',
            'summary': {
                'total': len(rows),
                'created': created_count,
                'skipped_duplicate': duplicate_count,
                'invalid': invalid_count,
            },
            'results': results,
        })
    except Exception as exc:
        return JsonResponse({
            'status': 'error',
            'message': str(exc),
        }, status=400)

@csrf_exempt
@require_http_methods(["POST"])
def add_form_api(request):
    try:
        data = json.loads(request.body)
        
        # 检查数据库是否有数据
        latest_record = FormsData.objects.all().order_by('-id').first()
        
        # 生成编号
        year_part = datetime.now().strftime('%y')
        if latest_record:
            if year_part == FormsData.get_number_year(latest_record):
                seq_number = int(FormsData.get_number_seq(latest_record)) + 1
                number = f'JCZX{year_part}{seq_number:04d}'
            else:
                number = f'JCZX{year_part}0001'
        else:
            number = f'JCZX{year_part}0001'
        
        # 创建数据对象
        form_data = FormsData(
            number=number,
            feedback_date=data.get('feedback_date'),
            feedback_department=data.get('feedback_department'),
            title=data.get('title'),
            author=data.get('author'),
            unit=data.get('unit'),
            accept_level=data.get('accept_level'),
            instruction_level=data.get('instruction_level'),
            remark=data.get('remark')
        )
        
        # 处理可选作者字段
        for i in range(2, 11):
            author_key = f'author_{i}'
            unit_key = f'unit_{i}'
            if author_key in data and unit_key in data:
                setattr(form_data, author_key, data[author_key])
                setattr(form_data, unit_key, data[unit_key])
        
        form_data.save()
        
        return JsonResponse({
            'status': 'success',
            'message': '数据保存成功',
            'number': number
        })
        
    except Exception as e:
        return JsonResponse({
            'status': 'error',
            'message': str(e)
        }, status=400)

@csrf_exempt
@require_http_methods(["PUT"])
def update_form_api(request, number):
    try:
        data = json.loads(request.body)
        form_data = get_object_or_404(FormsData, number=number)
        
        # 更新基础字段
        form_data.feedback_date = data.get('feedback_date')
        form_data.feedback_department = data.get('feedback_department')
        form_data.title = data.get('title')
        form_data.author = data.get('author')
        form_data.unit = data.get('unit')
        form_data.accept_level = data.get('accept_level')
        form_data.instruction_level = data.get('instruction_level')
        form_data.remark = data.get('remark')
        
        # 更新可选作者字段
        for i in range(2, 11):
            author_key = f'author_{i}'
            unit_key = f'unit_{i}'
            if author_key in data and unit_key in data:
                setattr(form_data, author_key, data[author_key])
                setattr(form_data, unit_key, data[unit_key])
        
        form_data.save()
        
        return JsonResponse({
            'status': 'success',
            'message': '数据更新成功'
        })
        
    except Exception as e:
        return JsonResponse({
            'status': 'error',
            'message': str(e)
        }, status=400)

@csrf_exempt
@require_http_methods(["DELETE"])
def delete_form_api(request, number):
    try:
        form_data = get_object_or_404(FormsData, number=number)
        
        # 检查记录是否已被删除
        if form_data.is_delete:
            return JsonResponse({
                'status': 'error',
                'message': '记录已被删除'
            }, status=400)
            
        # 软删除
        form_data.is_delete = True
        form_data.save()
        
        return JsonResponse({
            'status': 'success',
            'message': '数据删除成功',
            'number': number
        })
        
    except FormsData.DoesNotExist:
        return JsonResponse({
            'status': 'error',
            'message': '记录不存在'
        }, status=404)
        
    except Exception as e:
        return JsonResponse({
            'status': 'error',
            'message': str(e)
        }, status=500)

# TODO：完善统计分析逻辑
def statistics_form_template(request):
    selected_years = request.GET.getlist('selected_years')
    
    # 获取基础数据集
    data_list = FormsData.objects.filter(is_delete=False)
    
    # 按选中的年份筛选
    if selected_years:
        data_list = data_list.filter(feedback_date__year__in=map(int, selected_years))
    
    # 生成统计数据
    stats = {}
    # 添加总计统计
    total_stats = {
        'years': {},
        'accept_data': {
            'total': 0,
            'total_by_year': {},
            **{level: {'total': 0, 'years': {}} for level in STANDARD_ACCEPT_LEVELS}
        },
        'instruction_data': {
            'total': 0,
            'total_by_year': {},
            **{level: {'total': 0, 'years': {}} for level in STANDARD_INSTRUCTION_LEVELS}
        }
    }
    
    accept_levels = STANDARD_ACCEPT_LEVELS
    instruction_levels = STANDARD_INSTRUCTION_LEVELS
    
    for item in data_list:
        year = str(item.feedback_date.year)
        unit = item.unit  # 只取作者一的单位
        
        if not unit:
            continue
        
        # 初始化单位统计数据
        if unit not in stats:
            stats[unit] = {
                'total': 0,
                'years': {},
                'accept_data': {
                    'total': 0,
                    'total_by_year': {},
                    **{level: {'total': 0, 'years': {}} for level in accept_levels}
                },
                'instruction_data': {
                    'total': 0,
                    'total_by_year': {},
                    **{level: {'total': 0, 'years': {}} for level in instruction_levels}
                }
            }
        
        # 更新单位统计
        stats[unit]['total'] += 1
        stats[unit]['years'][year] = stats[unit]['years'].get(year, 0) + 1
        
        # 更新总计统计的年度数据
        total_stats['years'][year] = total_stats['years'].get(year, 0) + 1
        
        # 统计采纳情况
        accept_levels_in_item = [level for level in item.accept_level.split('、') if level]
        if accept_levels_in_item:
            # 更新单位统计
            stats[unit]['accept_data']['total'] += 1
            if year not in stats[unit]['accept_data']['total_by_year']:
                stats[unit]['accept_data']['total_by_year'][year] = 0
            stats[unit]['accept_data']['total_by_year'][year] += 1
            
            # 更新总计统计
            total_stats['accept_data']['total'] += 1
            if year not in total_stats['accept_data']['total_by_year']:
                total_stats['accept_data']['total_by_year'][year] = 0
            total_stats['accept_data']['total_by_year'][year] += 1
            
            for level in accept_levels_in_item:
                if level in accept_levels:
                    # 更新单位统计
                    stats[unit]['accept_data'][level]['total'] += 1
                    if year not in stats[unit]['accept_data'][level]['years']:
                        stats[unit]['accept_data'][level]['years'][year] = 0
                    stats[unit]['accept_data'][level]['years'][year] += 1
                    
                    # 更新总计统计
                    total_stats['accept_data'][level]['total'] += 1
                    if year not in total_stats['accept_data'][level]['years']:
                        total_stats['accept_data'][level]['years'][year] = 0
                    total_stats['accept_data'][level]['years'][year] += 1
        
        # 统计批示情况
        instruction_levels_in_item = [level for level in item.instruction_level.split('、') if level]
        if instruction_levels_in_item:
            # 更新单位统计
            stats[unit]['instruction_data']['total'] += 1
            if year not in stats[unit]['instruction_data']['total_by_year']:
                stats[unit]['instruction_data']['total_by_year'][year] = 0
            stats[unit]['instruction_data']['total_by_year'][year] += 1
            
            # 更新总计统计
            total_stats['instruction_data']['total'] += 1
            if year not in total_stats['instruction_data']['total_by_year']:
                total_stats['instruction_data']['total_by_year'][year] = 0
            total_stats['instruction_data']['total_by_year'][year] += 1
            
            for level in instruction_levels_in_item:
                if level in instruction_levels:
                    # 更新单位统计
                    stats[unit]['instruction_data'][level]['total'] += 1
                    if year not in stats[unit]['instruction_data'][level]['years']:
                        stats[unit]['instruction_data'][level]['years'][year] = 0
                    stats[unit]['instruction_data'][level]['years'][year] += 1
                    
                    # 更新总计统计
                    total_stats['instruction_data'][level]['total'] += 1
                    if year not in total_stats['instruction_data'][level]['years']:
                        total_stats['instruction_data'][level]['years'][year] = 0
                    total_stats['instruction_data'][level]['years'][year] += 1
    
    # 获取所有年份列表
    years = sorted(list(set(str(item.feedback_date.year) for item in data_list)))
    
    return render(request, 'forms/statistics.html', {
        'stats': stats,
        'total_stats': total_stats,  # 添加总计统计数据
        'years': years,
        'selected_years': selected_years,
        'standard_years': STANDARD_YEARS,
        'accept_levels': accept_levels,
        'instruction_levels': instruction_levels
    })