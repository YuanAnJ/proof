from django.db import models

# Create your models here.
class FormsData(models.Model):
    number = models.CharField('信息编号',max_length=255,default='')
    feedback_date = models.DateField('反馈日期')
    feedback_department = models.CharField('反馈部门',max_length=255,default='')
    title = models.CharField('报告题目',max_length=255,default='')
    author = models.CharField('作者一',max_length=255,default='')
    unit = models.CharField('作者一所在单位',max_length=255,default='')
    author_2 = models.CharField('作者二',max_length=255,default='')
    unit_2 = models.CharField('作者二所在单位',max_length=255,default='')
    author_3 = models.CharField('作者三',max_length=255,default='')
    unit_3 = models.CharField('作者三所在单位',max_length=255,default='')
    author_4 = models.CharField('作者四',max_length=255,default='')
    unit_4 = models.CharField('作者四所在单位',max_length=255,default='')
    author_5 = models.CharField('作者五',max_length=255,default='')
    unit_5 = models.CharField('作者五所在单位',max_length=255,default='')
    author_6 = models.CharField('作者六',max_length=255,default='')
    unit_6 = models.CharField('作者六所在单位',max_length=255,default='')
    author_7 = models.CharField('作者七',max_length=255,default='')
    unit_7 = models.CharField('作者七所在单位',max_length=255,default='')
    author_8 = models.CharField('作者八',max_length=255,default='')
    unit_8 = models.CharField('作者八所在单位',max_length=255,default='')
    author_9 = models.CharField('作者九',max_length=255,default='')
    unit_9 = models.CharField('作者九所在单位',max_length=255,default='')
    author_10 = models.CharField('作者十',max_length=255,default='')
    unit_10 = models.CharField('作者十所在单位',max_length=255,default='')
    accept_level = models.CharField('采纳级别',max_length=255,default='')
    other_accept_level = models.CharField('其他采纳级别',max_length=255,default='')
    instruction_level = models.CharField('批示级别',max_length=255,default='')
    other_instruction_level=models.CharField('其他批示级别',max_length=255,default='')
    remark = models.CharField('备注',max_length=255,default='')
    is_delete = models.BooleanField('是否删除',default=False)
    
    def get_number_year(self):
        """获取编号的年份信息"""
        if len(self.number) >= 6:
            return self.number[4:6]
        return ''
    
    def get_number_seq(self):
        """获取编号的序列号"""
        if len(self.number) >= 10:
            return self.number[6:10]
        return ''

    def get_author_info(self):
        """返回所有作者及其单位信息"""
        authors_info = []
        for i in range(1, 11):
            author = getattr(self, f'author_{i}' if i > 1 else 'author')
            unit = getattr(self, f'unit_{i}' if i > 1 else 'unit')
            if author and unit:  # 只返回非空的作者和单位信息
                authors_info.append({
                    'unit': unit,
                    'author': author
                })
        return authors_info
    

