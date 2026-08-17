# -*- coding: utf-8 -*-
"""
生成少量演示原始订单 Excel（仅用于随仓库附带的样例 / 回归测试）。
· 悟空：列名含「商户应收 / 取车城市」，平台自动识别为 悟空。
· 滴滴：列名含「车型信息 / 车辆信息 / 订单金额 / 取车门店」，自动识别为 滴滴。
运行：
    python make_demo_data.py
输出：
    ../sample-data/悟空_取车订单_demo.xlsx
    ../sample-data/滴滴_取车订单_demo.xlsx
随后运行 generate_dashboard.py 即可由这些文件生成归一化 data.json。
"""
import os
import random
from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLE = os.path.join(HERE, '..', 'sample-data')
os.makedirs(SAMPLE, exist_ok=True)

random.seed(20260817)

# 悟空：城市直接写城市名（城市映射为 null）
WK_CITIES = ['北京', '重庆', '广州', '深圳', '西安']
WK_MODELS = ['广汽埃安 AION S', '北京 EU5', '丰田bZ3', '风神SKY EV01', '帝豪新能源']
WK_STATUSES = [('已完成', '待违章处理'), ('待取车', '待用车'), ('待还车', '用车中'), ('已取消', '已取消')]

# 滴滴：用门店名（命中 滴滴 city_maps 映射成城市）
DD_STORES = {
    '快快租车-北京店': '北京', '快快租车-深圳店': '深圳',
    '快快租车-广州店': '广州', '快快租车-西安店': '西安',
}
DD_MODELS = ['AION S', '北京EU5 PLUS', '雷凌', '秦新能源', '零跑B01']
DD_STATUSES = [('已完成', '已还车'), ('待取车', '待取车'), ('待还车', '用车中'), ('已取消', '已取消')]


def rnd_plate(city_prefix):
    letters = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    return city_prefix + ''.join(random.choice(letters + '0123456789') for _ in range(5))


def make_wukong(n):
    wb = Workbook()
    ws = wb.active
    ws.title = '订单'
    ws.append(['订单号', '下单时间', '预计取车时间', '实际取车时间', '预计还车时间',
               '实际还车时间', '租期', '车型', '取车城市', '商户应收', '订单状态', '车牌号'])
    base = 2976800000000
    for i in range(n):
        city = random.choice(WK_CITIES)
        model = random.choice(WK_MODELS)
        status, raw = random.choice(WK_STATUSES)
        month = random.choice(['2026-07', '2026-08'])
        day = random.randint(1, 20)
        pickup = f'{month}-{day:02d}'
        rent = random.choice([1, 1, 2, 3, 1])
        ret = f'{month}-{day+rent:02d}'
        amt = random.choice([98, 128, 156, 188, 211, 256, 299, 168])
        # 已取消订单无实际取还车
        if status == '已取消':
            apick = aret = ''
        else:
            apick = pickup
            aret = ret if status in ('已完成', '待还车') else ''
        ws.append([str(base + i), pickup[:7] + '-' + pickup[8:], pickup, apick, ret, aret,
                   f'{rent}天', model, city, amt, raw, rnd_plate('京' if city == '北京' else '渝' if city == '重庆' else '粤')])
    path = os.path.join(SAMPLE, '悟空_取车订单_demo.xlsx')
    wb.save(path)
    return path


def make_didi(n):
    wb = Workbook()
    ws = wb.active
    ws.title = '订单'
    ws.append(['订单编号', '下单时间', '取车时间', '实际取车时间', '还车时间',
               '实际还车时间', '租期', '车型信息', '取车门店', '订单金额', '订单状态', '车辆信息'])
    base = 883200000000
    for i in range(n):
        store = random.choice(list(DD_STORES.keys()))
        city = DD_STORES[store]
        model = random.choice(DD_MODELS)
        status, raw = random.choice(DD_STATUSES)
        month = random.choice(['2026-07', '2026-08'])
        day = random.randint(1, 18)
        pickup = f'{month}-{day:02d}'
        rent = random.choice([1, 2, 1, 3])
        ret = f'{month}-{day+rent:02d}'
        amt = random.choice([110, 145, 178, 199, 228, 268])
        if status == '已取消':
            apick = aret = ''
        else:
            apick = pickup
            aret = ret if status in ('已完成', '待还车') else ''
        ws.append([str(base + i), pickup, pickup, apick, ret, aret,
                   f'{rent}天', model, store, amt, raw, rnd_plate('京' if city == '北京' else '粤')])
    path = os.path.join(SAMPLE, '滴滴_取车订单_demo.xlsx')
    wb.save(path)
    return path


if __name__ == '__main__':
    p1 = make_wukong(18)
    p2 = make_didi(14)
    print('已生成演示数据：')
    print(' ', p1)
    print(' ', p2)
