from django import template

register = template.Library()

@register.filter(name='get_item')
def get_item(dictionary, key):
    """
    获取字典中的指定键值
    用法: {{ dictionary|get_item:key }}
    """
    if dictionary is None:
        return None
    return dictionary.get(key)
