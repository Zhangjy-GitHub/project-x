from __future__ import unicode_literals


def convert(amount):
    num = amount.split('.', 1)
    units = ['', '万', '亿']
    nums = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    decimal_label = ['角', '分']
    small_int_label = ['', '拾', '佰', '仟']
    int_part = num[0]
    decimal_part = '0'
    if len(num) == 2:
        decimal_part = num[1]
    res = []
    if decimal_part != '0':
        res.append(''.join([nums[int(x)] + y for x, y in zip(decimal_part, decimal_label) if x != '0']))
    else:
        res.append('整')
    if int_part != '0':
        res.append('圆')
        while int_part:
            small_int_part, int_part = int_part[-4:], int_part[:-4]
            tmp = ''.join([nums[int(x)] + (y if x != '0' else '') for x, y \
                           in list(zip(small_int_part[::-1], small_int_label))[::-1]])
            tmp = tmp.rstrip('零').replace('零零零', '零').replace('零零', '零')
            unit = units.pop(0)
            if tmp:
                tmp += unit
                res.append(tmp)
    return ''.join(res[::-1])


if __name__ == '__main__':
    print(convert('6000.16'))
