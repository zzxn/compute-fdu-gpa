#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author: zzxn

import sys

import xlrd


def get_monospaced_str_list(str_list):
    max_str_len = 0
    for s in str_list:
        max_str_len = max_str_len if max_str_len > len(s) else len(s)
    return [str_b_to_q(s + '　' * (max_str_len - len(s))) for s in str_list]


def str_b_to_q(b_str):
    q_str = ''
    for uchar in b_str:
        inside_code = ord(uchar)
        if inside_code == 32:
            inside_code = 12288
        elif 32 <= inside_code <= 126:
            inside_code += 65248

        q_str += chr(inside_code)
    return q_str


def print_course_info(excel_file_path, detail=True, color=False):
    # open excel table and sheet 0
    table = xlrd.open_workbook(excel_file_path)
    sheet = table.sheet_by_index(0)

    # print total course number
    print('*****************************************************************')
    print('文件名：　', excel_file_path)
    print('总课程数：', sheet.nrows)
    print('* 注：务必手动删除P和NP等不计算绩点的课程')

    # get course names, credits, levels and points
    course_names = sheet.col_values(3)
    course_credits = sheet.col_values(4)
    course_levels = sheet.col_values(5)
    course_points = sheet.col_values(6)

    # make each char in course_names monospaced
    monospaced_course_names = get_monospaced_str_list(course_names)

    # print course information and compute total credits and weighted points
    total_credits = 0.0
    weighted_total_points = 0.0
    for i in range(sheet.nrows):
        total_credits += course_credits[i]
        weighted_total_points += course_credits[i] * course_points[i]

    gpa = weighted_total_points / total_credits

    if detail:
        print('-----------------------------------------------------------------')
        for i in range(sheet.nrows):
            gpa_without_this_course = (weighted_total_points - course_credits[i] * course_points[i]) \
                                      / (total_credits - course_credits[i])
            impact = gpa - gpa_without_this_course

            if color:
                if 0 <= impact < 0.01:
                    color_str = '\033[0;36m'
                elif impact >= 0.01:
                    color_str = '\033[0;32m'
                elif -0.01 < impact < 0:
                    color_str = '\033[0;33m'
                else:
                    color_str = '\033[0;31m'
                print(i + 1, '\t', monospaced_course_names[i], ' 学分：', course_credits[i],
                      '\t等级：', course_levels[i], '\t绩点：', course_points[i],
                      '\t影响：', color_str, format(impact, '+.3f'), '\033[0m')
            else:
                print(i + 1, '\t', monospaced_course_names[i], ' 学分：', course_credits[i],
                      '\t等级：', course_levels[i], '\t绩点：', course_points[i],
                      '\t影响：', format(impact, '+.3f'))

    # compute and print statistics
    print('-----------------------------------------------------------------')
    print('总学分：　　　', total_credits)
    print('加权总绩点：　', weighted_total_points)
    print('平均绩点：　　', format(weighted_total_points / total_credits, '.3f'))
    print('*****************************************************************')


def print_help():
    print("用法: [python]", sys.argv[0], 'filename', '[-s]', '[-c]')
    print('filename 是包含成绩表格的Excel文件，成绩表格请到http://jwfw.fudan.edu.cn/处复制到Excel中，然后保存')
    print('* 注：需要手动删除不需要参与计算的课程（包括以P/NP计的课程）')
    print('-s 是可选参数，指定是否省略打印详细信息，默认会打印详细信息')
    print('-c 是可选参数，指定输出是否是彩色的，默认是非彩色的，目前只支持Bash')


if __name__ == '__main__':
    if len(sys.argv) < 2 or sys.argv.count('-h') != 0 or sys.argv.count('--help') != 0:
        print_help()
    else:
        print_course_info(sys.argv[1], sys.argv.count('-s') == 0, sys.argv.count('-c') != 0)
