from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse, JsonResponse
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.db.models import Q
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
import json
from .models import FormsData  # 确保导入你的数据模型
from datetime import datetime
import pandas as pd
from forms.utils.gd import generate_doc

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

# TODO：完善搜索逻辑，分页逻辑，跳转逻辑
def query_form(request):
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

    response = HttpResponse(open('media/证明文件.docx', 'rb'), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = f'attachment; filename="{file_name}_Proof.doc"'
    return response

def export_excel_api(request):
    # 获取所有未删除的数据
    data = FormsData.objects.filter(is_delete=False)
    
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
def statistics_form(request):
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