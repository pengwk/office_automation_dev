# -*- coding: utf-8 -*-

import os

from win32com.client.gencache import EnsureDispatch

import pythoncom


def main():

    try:
        excel = EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.Workbooks.Open("error_path.xls")
    except pythoncom.com_error as _error:
        if _error.excepinfo is None:
            excepinfo = [None]*6
        else:
            excepinfo = _error.excepinfo
            hresult_readable = u"""
    hresult状态码：{0}
    hresult状态码信息：{1}"""
            detail_readable = u"""
    错误码：  {0}
    错误源：  {1}
    错误信息：{2}
    帮助文档: {3}
    帮助ID：  {4}
    错误码：  {5}
                    """
        args_readable = u"""参数错误，参数所在位置：{0}"""
        print hresult_readable.format(_error.hresult, _error.strerror.decode('gbk'))
        print detail_readable.format(*excepinfo)
        if _error.argerror > 0:
            print args_readable.format(_error.argerror)
    return None


if __name__ == '__main__':
    main()