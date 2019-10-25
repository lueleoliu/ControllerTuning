# -*- coding: utf-8 -*-

import GW_functions as gw
import os
import shutil
import psutil
import threading
import math


def version_check(prj_path):
    version_info = gw.get_info(prj_path, 'VERSION')
    try:
        if '4.7' in version_info:
            version = '4.7'
        elif '4.6' in version_info:
            version = '4.6'
        else:
            raise Exception('不支持的Bladed版本')
        return version
    except Exception as e:
        return e


class runbat(threading.Thread):
    def __init__(self, path, type):
        threading.Thread.__init__(self)
        self.path = path
        self.type = type

    def run(self):
        gw.run_bat(self.path, self.type)


class pidcal(threading.Thread):
    def __init__(self, path):
        threading.Thread.__init__(self)
        self.path = path

    def run(self):
        gw.pid_cal(self.path)


def terminate(file_dir, app, msg):
    pl = psutil.pids()
    proc = []
    for pid in pl:
        if psutil.Process(pid).name() == app:
            print(pid)
            try:
                proc.append(psutil.Process(pid))
                for p in proc:
                    p.terminate()
                gone, alive = psutil.wait_procs(proc, timeout=3)
                for p in alive:
                    p.kill()
            finally:
                gw.logging(file_dir, msg)
                raise Exception(msg)


def pre_processing(files_dir):
    prj_dir = os.path.join(files_dir, 'Model.prj')
    vs = version_check(prj_dir)

    if isinstance(vs, Exception):
        return vs

    if vs == '4.6':
        config_dir = os.path.join(files_dir, 'config_4.6.txt')
    elif vs == '4.7':
        config_dir = os.path.join(files_dir, 'config_4.7.txt')
    else:
        config_dir = vs

    with open(config_dir, 'r') as file_to_read:
        lines = file_to_read.read()

    dir = lines
    model_code = ['@echo off\n',
                  'cd %~dp0\n',
                  'setlocal EnableDelayedExpansion\n',
                  '''for /f "delims=" %%i in ('"dir /aa/s/b/on *.prj"') do (\n''',
                  'set file=%%~fi\n',
                  'set file=!file:/=/!\n',
                  ')\n',
                  '\n',
                  'set apath=%~dp0\n',
                  '\n',
                  r'"' + dir + r'Bladed_m72.exe" -Prj !file! -RunDir !apath! -ResultPath !apath!']
    cal_code = ['@echo off\n',
                'cd %~dp0\n',
                r'"' + os.path.join(dir, 'dtbladed.exe') + r'"']
    linmod_code = ['@echo off\n',
                   'cd %~dp0\n',
                   r'"' + dir + r'dtlinmod.exe"']

    gw.mkdir(files_dir, 'Temp')
    exctrl_dir = gw.mkdir(files_dir, 'Exctrl')
    model_dir = gw.mkbat(files_dir, 'Model', model_code)
    gw.mkbat(files_dir, 'Performance', cal_code)
    gw.mkbat(files_dir, 'Campbell', cal_code)
    gw.mkbat(files_dir, 'Linear', cal_code)
    gw.mkbat(files_dir, 'LinearModel', linmod_code)

    shutil.copy(prj_dir, os.path.join(model_dir, 'Model.prj'))

    if vs == '4.7':
        prj_m_dir = os.path.join(files_dir, 'Model_m.prj')
        model_m_dir = gw.mkbat(files_dir, 'Model_m', model_code)
        shutil.copy(prj_m_dir, os.path.join(model_m_dir, 'Model.prj'))
    else:
        pass

    dll_dirs = gw.get_typefile(files_dir, '.dll', 'Discon')

    xml_dirs = gw.get_typefile(files_dir, '.xml', 'Parameters')

    if len(dll_dirs) > 1:
        gw.logging(files_dir, ' ERROR : 控制器dll文件过多！')
        raise Exception('控制器dll文件过多！')

    elif not dll_dirs:
        gw.logging(files_dir, ' ERROR : 没有控制器dll文件！')
        raise Exception('没有控制器dll文件！')

    if len(xml_dirs) > 1:
        gw.logging(files_dir, ' ERROR : 控制器xml文件过多！')
        raise Exception('控制器xml文件过多！')

    elif not xml_dirs:
        gw.logging(files_dir, ' ERROR : 没有控制器xml文件！')
        raise Exception('没有控制器xml文件')

    shutil.copy(dll_dirs[0], exctrl_dir)
    shutil.copy(xml_dirs[0], exctrl_dir)

    project_dir = gw.get_typefile(model_dir, '.prj')

    dll_file_path = gw.get_typefile(exctrl_dir, '.dll')
    xml_file_path = gw.get_typefile(exctrl_dir, '.xml')

    gw.change_xml(project_dir[0], 'ExternalController', 'Filepath', dll_file_path[0])
    gw.change_xml(project_dir[0], 'ExternalController', 'AdditionalParameters', 'READ ' + xml_file_path[0])

    gw.change_info(project_dir[0], 'CALCULATION', '2')

    t_1 = runbat(files_dir, 'Model')
    t_1.start()
    t_1.join(timeout=50)
    terminate(files_dir, 'Bladed_m72.exe', ' Error: Model open failed')

    in_dir = gw.get_typefile(model_dir, '.in')[0]
    eigen_b = gw.catch_block(in_dir, 'EIGENB')
    eigen_b.append('\n')
    eigen_t = gw.catch_block(in_dir, 'EIGENT')
    eigen_t.append('\n')

    gw.delete_block(project_dir[0], 'RMODE')
    gw.delete_info(project_dir[0], '0RMASS')

    gw.change_info(project_dir[0], 'CALCULATION', '10')

    t_2 = runbat(files_dir, 'Model')
    t_2.start()
    t_2.join(timeout=50)
    terminate(files_dir, 'Bladed_m72.exe', ' Error: Model open failed')

    gw.add_block(in_dir, 'RCON', eigen_b)
    gw.add_block(in_dir, 'RCON', eigen_t)

    gw.logging(files_dir, ' 初始in文件生成成功')
    print('初始in文件生成成功')

    ori_in_dir = os.path.join(files_dir, 'Model')
    ori_in_file_path = gw.get_typefile(ori_in_dir, '.in')
    current_in_dir = os.path.join(files_dir, 'Performance')
    current_in_file_path = ori_in_file_path[0].replace(ori_in_dir, current_in_dir)
    shutil.copyfile(ori_in_file_path[0], current_in_file_path)

    gw.change_info(current_in_file_path, 'CALCN', '5')
    gw.change_info(current_in_file_path, 'PATH', current_in_dir)
    gw.change_info(current_in_file_path, 'RUNNAME', 'pcoeffs')
    gw.change_info(current_in_file_path, 'OPTNS', '0')
    gw.change_block(current_in_file_path, 'PCOEFF', 'PITCH', '-0.03490660')
    gw.change_block(current_in_file_path, 'PCOEFF', 'PITCH_END', '0.03490660')
    gw.change_block(current_in_file_path, 'PCOEFF', 'PITCH_STEP', '0.00872665')

    print('开始计算最小桨距角')
    gw.run_bat(files_dir, 'Performance')
    gw.logging(files_dir, ' 最小桨距角计算完成')
    print('最小桨距角计算完成')

    ori_in_dir = os.path.join(files_dir, 'Model')
    ori_in_file_path = gw.get_typefile(ori_in_dir, '.in')

    gw.change_block(ori_in_file_path[0], 'CONTROL', 'GAIN_TSR', str(get_optmodegain(files_dir)))
    gw.change_block(ori_in_file_path[0], 'CONTROL', 'PITMIN', get_cpinfo(files_dir)[2])

    gw.logging(files_dir, ' 修正in文件完成')
    print('修正in文件完成')


def model_correction(files_dir):
    prj_dir = os.path.join(files_dir, 'Model.prj')
    vs = version_check(prj_dir)

    if isinstance(vs, Exception):
        return vs

    if vs == '4.6':
        pass
    elif vs == '4.7':
        model_m_dir = os.path.join(files_dir, 'Model_m')

        if os.path.exists(model_m_dir):
            exctrl_dir = os.path.join(files_dir, 'Exctrl')

            project_dir = gw.get_typefile(model_m_dir, '.prj')

            dll_file_path = gw.get_typefile(exctrl_dir, '.dll')
            xml_file_path = gw.get_typefile(exctrl_dir, '.xml')

            gw.change_xml(project_dir[0], 'ExternalController', 'Filepath', dll_file_path[0])
            gw.change_xml(project_dir[0], 'ExternalController', 'AdditionalParameters', 'READ ' + xml_file_path[0])

            gw.change_info(project_dir[0], 'CALCULATION', '2')
            gw.run_bat(files_dir, 'Model')
            in_m_dir = gw.get_typefile(model_m_dir, '.in')[0]
            eigen_b = gw.catch_block(in_m_dir, 'EIGENB')
            eigen_b.append('\n')
            eigen_t = gw.catch_block(in_m_dir, 'EIGENT')
            eigen_t.append('\n')

            gw.delete_block(project_dir[0], 'RMODE')
            gw.delete_info(project_dir[0], '0RMASS')

            gw.change_info(project_dir[0], 'CALCULATION', '10')
            gw.run_bat(files_dir, 'Model')

            gw.add_block(in_m_dir, 'RCON', eigen_b)
            gw.add_block(in_m_dir, 'RCON', eigen_t)

            model_dir = os.path.join(files_dir, 'Model')
            in_dir = gw.get_typefile(model_dir, '.in')[0]

            kopt = gw.get_block(in_dir, 'CONTROL', 'GAIN_TSR')
            pit_min = gw.get_block(in_dir, 'CONTROL', 'PITMIN')

            gw.change_block(in_m_dir, 'CONTROL', 'GAIN_TSR', str(kopt))
            gw.change_block(in_m_dir, 'CONTROL', 'PITMIN', str(pit_min))

            gw.logging(files_dir, ' 分段模型in文件生成成功')
            print('分段模型in文件生成成功')

        else:
            pass

    else:
        pass


def get_cpinfo(root_dir):
    prj_dir = os.path.join(root_dir, 'Model.prj')
    vs = version_check(prj_dir)

    cpinfo = []
    data_path = []
    cp_path = os.path.join(root_dir, 'Performance')
    if vs == '4.6':
        data_path = gw.get_typefile(cp_path, '.%37')
    elif vs == '4.7':
        data_path = gw.get_typefile(cp_path, '.%55')
    else:
        pass

    with open(data_path[0], 'r') as msg:
        lines = msg.readlines()
        for j in range(len(lines)):
            if 'ULOADS' in lines[j]:
                cpinfo.append(gw.do_split(lines[j], '  ', 1))     # 提取Cp
                cpinfo.append(gw.do_split(lines[j], '  ', 2))     # 提取λ
            elif 'MAXTIME'in lines[j]:
                cpinfo.append(gw.do_split(lines[j], '  ', 1))     # 提取最小桨距角
    return cpinfo


def get_optmodegain(root_dir):
    global logfile

    cpinfo = get_cpinfo(root_dir)
    cp_max = float(cpinfo[0].strip())
    lamda = float(cpinfo[1].strip())

    model_dir = os.path.join(root_dir, 'Model')
    project_dir = gw.get_typefile(model_dir, '.prj')
    radius = 0.5 * gw.get_block(project_dir[0], 'RCON', 'DIAM')

    pho = gw.get_block(project_dir[0], 'CONSTANTS', 'RHO')

    k_opt = math.pi * pho * math.pow(radius, 5) * cp_max / (2 * math.pow(lamda, 3))
    gw.logging(root_dir, ' Kopt计算完成')
    print('Kopt计算完成')
    return int(k_opt)


if __name__ == '__main__':
    path = os.path.join(os.path.abspath('..'), 'files')
    pre_processing(path)

    gw.gen_campbell(path)
    gw.gen_linear_model(path)
    gw.get_wt_basic_info(path)

    t_3 = pidcal(path)
    t_3.start()
    t_3.join(timeout=10800)
    terminate(path, 'GenFile.exe', ' Error: Linearization result was incorrect, check model please')

    gw.logging(path, ' PID优化计算完成')
    gw.get_result(path)
    gw.print_pid_to_xml(path)
    gw.logging(path, ' PID信息修改完毕')
    gw.print_filters_to_xml(path)
    gw.logging(path, ' 滤波器信息修改完毕')

    model_correction(path)
