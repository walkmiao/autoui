# This is a sample Python script.
# !/user/bin/env python3
# -*- coding: utf-8 -*-
import logging
import time
from collections import namedtuple
from logging.handlers import TimedRotatingFileHandler

import psutil

import dlg
from para import BaoYangQuery,BaoYangEdit
from pywinauto import Application, mouse
from pywinauto.controls.common_controls import DateTimePickerWrapper
from pywinauto.keyboard import send_keys
from pywinauto import  mouse
from pywinauto.controls.uia_controls import ListViewWrapper

from data import op_map
from dlg import combox_select,save_and_exit
from utils import get_pid, create_logger


class Handle:
    def __init__(self, work_name):
        """

        @param main_window: 程序的主窗体
        """
        self.run_process_name = 'TDSV2.FAWJF.exe'
        self.start_process_name = 'E:\TDS助手外网\SUT.Client.Upgrade.exe'
        self._cur_window = None
        self._oper_window = None
        self._window = None
        self.app = None
        self._logger = None
        self.op_map = op_map
        self.inited = False
        self.work_name = work_name
        self.set_loggger()
        self.login()

    def set_loggger(self):
        self._logger = create_logger(self.work_name)

    def init_window(self, *, tree_item_title, node, print_control=False):
        """
        初始化操作窗口
        @return:
        """
        op_para = self.get_op_para(tree_item_title, node)
        if not op_para:
            self._logger.error('获取业务{}-{}参数失败'.format(tree_item_title, node))
            return
        title = op_para.get('op_title')
        if not title:
            self._logger.error('获取业务{}-{}窗格代码失败'.format(tree_item_title, node))
            return
        op_w = self.get_op_window(tree_item_title=tree_item_title, op_title=title, node_name=node,
                                  control_type="Window")
        op_w.wait('ready', timeout=10)
        self.inited = True
        if print_control and op_w:
            op_w.print_control_identifiers()

    def get_op_para(self, yewu_name, yewu_item_name):
        yewu = self.op_map.get(yewu_name)
        if yewu:
            return yewu.get(yewu_item_name)

    def login(self):
        try:

            pid = get_pid(self.run_process_name)
            if not pid:
                self._logger.warning("程序{}还未运行".format(self.run_process_name))
                up_app = Application(backend='uia').start(self.start_process_name)
                # up_app.top_window().wait_not('exists')
                time.sleep(3)
                pid = get_pid(self.run_process_name)
                self.app = Application(backend='uia').connect(process=pid)
                login_w = self.app['登陆界面']
                login_w.wait('ready', timeout=10)
                login_w.child_window(auto_id="txtUserName", control_type="Edit").set_edit_text('FQD0209')
                login_w.child_window(auto_id="txtPassWord", control_type="Edit").set_edit_text('Piccsh@00007')
                login_w.child_window(title="新解放正式环境", auto_id="lbLogin", control_type="Text").click_input()
                self._logger.info('登陆成功')
            else:
                self._logger.info("进程【{0}】已经在运行，pid:【{1}】".format(self.run_process_name, pid))
                self.app = Application(backend='uia').connect(process=pid)
        except Exception as e:
            self._logger.exception(e)

    def get_cur_window(self, *, title, control_type='TreeItem'):
        """
        设置当前的操作子窗口
        @param title: 窗口名次
        @param control_type: 控制类型
        @return: 子窗口
        """
        self._window = self.app['主界面']
        self._cur_window = self._window.child_window(title=title, control_type=control_type)
        if self._cur_window.exists():
            self._logger.info('获取到当前节点窗口:{}'.format(title))
        else:
            self._logger.exception('还未获取到当前节点窗口:{}'.format(title))

    def get_op_window(self, *, tree_item_title, op_title, node_name, control_type='Window'):
        """
        先获取菜单窗口，然后获取操作子窗口
        @param tree_item_title: 左侧菜单选择窗口名称
        @param op_title:右侧操作窗口名称
        @param node_name: 左侧选择的节点名称
        @param control_type: 控制类型
        @return: 当前操作窗口
        """
        self.get_cur_window(title=tree_item_title)
        self.expand_tree_item(node_name=node_name)
        self._oper_window = self._window.child_window(title=op_title, control_type=control_type)
        if self._oper_window.exists():
            self._logger.info('获取到当前操作窗口:{}'.format(op_title))
            return self._oper_window
        else:
            self._logger.exception('还未获取到当前操作窗口:{}'.format(op_title))

    def expand_tree_item(self, *, node_name):
        """
        根据窗口获取节点，展开-选择-双击
        @param window: _cur_window
        @param node_name: 节点名称
        @return:
        """
        self._cur_window.expand()
        node = self._cur_window.get_child(node_name)
        if not node:
            self._logger.exception('节点:{}不存在'.format(node_name))
        self._logger.debug('获取到节点【{}】,开始选择并双击'.format(node_name))
        node.select()
        node.double_click_input()

    def set_date_pane(self, *, pane, date_str):
        """

        @param pane: 日期pane
        @param date_str:
        @return:
        """
        picker_wrapper = DateTimePickerWrapper(pane.handle)
        try:
            year, month, day = date_str.split('-')
        except ValueError:
            self._logger.exception('日期字符串格式必须为list或者tuple')
        else:
            picker_wrapper.set_time(year=int(year), month=int(month), day=int(day))

    def close_op_window(self, key='{VK_F12}'):
        send_keys(key)


class BackupBusiness(Handle):
    def __init__(self, main_window):
        super().__init__(main_window)
        self._back_window = main_window

    def order_browse(self):
        self.get_cur_window()
        service_window = self._window.child_window(title="服务站业务", control_type="TreeItem")
        service_window.expand()
        query = service_window.get_child(u'维护服务站订单')
        # 选择当前节点
        query.select()
        # print(dir(query))
        # 双击当前节点
        query.double_click_input()
        # service_window.print_control_identifiers()
        # 找到额外的窗格
        service_order_window = self._window.child_window(title='W_QueryMessage', control_type="Window")


# 索赔业务

class SuoPeiService(Handle):
    def __init__(self):
        super().__init__(work_name='索赔业务')
        self.works = ['excel', 'query']

    def export_xlsx(self, filename):
        if not self.inited:
            self._logger.error('还未初始化过操作窗口')
            return
        export_btn = self._oper_window.child_window(title='导出Excel')
        export_btn.click_input()
        export_w = self.app['主界面']
        export_w.set_focus()
        down_btn = export_w.child_window(title="确认", auto_id="button1", control_type="Button")
        if down_btn.exists():
            down_btn.click()
        else:
            self._logger.info('未发现弹窗，退出')
        # export_w.print_control_identifiers()
        # export_w.child_window(title="地址: 文档", auto_id="1001", control_type="ToolBar").set_edit_text('E:/')
        export_w.child_window(title="文件名:", auto_id="1001", control_type="Edit").set_edit_text(filename)
        save_btn = export_w.child_window(title="保存(S)", auto_id="1", control_type="Button")
        save_btn.click()
        no_btn = export_w.child_window(title='否(N)')
        no_btn.click()
        self._logger.info('导出excel:{} 成功!'.format(filename))

    def query(self, create_dates, repair_dates, vin, sel):
        """
        维护索赔单
        @param create_dates:
        @param repair_dates:
        @param vin:
        @param sel:
        @return:
        """
        if len(create_dates) == 2:
            start1, stop1 = create_dates
            if start1 and stop1:
                tibao_start_date = self._oper_window.child_window(auto_id="dtUpdBegin")
                tibao_start_date.type_keys(start1)
                tibao_stop_date = self._oper_window.child_window(auto_id="dtUpdEnd")
                tibao_stop_date.type_keys(stop1)
            self._logger.debug('设置提报开始日期:{},结束日期:{}'.format(start1, stop1))
        else:
            self._logger.debug('提报日期为空，跳过设置')
        if len(repair_dates) == 2:
            start2, stop2 = repair_dates
            if start2 and stop2:
                repair_start = self._oper_window.child_window(auto_id="dtcSendDateBegin")
                repair_start.type_keys(start2)
                repair_stop = self._oper_window.child_window(auto_id="dtcSendTimeEnd")
                repair_stop.type_keys(stop2)
            self._logger.debug('设置提报开始日期:{},结束日期:{}'.format(start2, stop2))
        else:
            self._logger.debug('送修日期为空，跳过设置')
        # self._oper_window.print_control_identifiers()

        danju_status = self._oper_window.child_window(title="单据状态", auto_id="cobCombobox")
        combox_select(sel, danju_status)
        # service_domain = self._oper_window.child_window(title="*服务域", auto_id="cobCombobox", control_type="ComboBox")
        vin_code = self._oper_window.child_window(auto_id="txtVVin")
        if vin:
            self._logger.debug('设置vin:{}'.format(vin))
            vin_code.type_keys(vin)
        query_btn = self._oper_window.child_window(title="查询(Q)", auto_id="button1")
        query_btn.click()

    def query_zdspspt(self, *, oper_dates, order, vin, sel):
        """
        维护重大索赔审批单
        @param oper_dates:
        @param order:
        @param vin:
        @param sel:
        @return:
        """
        # 操作日期设置
        if len(oper_dates) == 2:
            start1, stop1 = oper_dates
            if start1 and stop1:
                start_date = self._oper_window.child_window(auto_id="dt_begin")
                start_date.type_keys(start1)
                stop_date = self._oper_window.child_window(auto_id="dt_end")
                stop_date.type_keys(stop1)
            self._logger.debug('设置操作开始日期:{},结束日期:{}'.format(start1, stop1))
        else:
            self._logger.debug('操作日期为空，跳过设置')

        # 状态选择
        danju_status = self._oper_window.child_window(auto_id="cbcVFinishStateQ")
        combox_select(sel, danju_status)

        # vin
        vin_code = self._oper_window.child_window(auto_id="txtVVIN")
        if vin:
            self._logger.debug('设置vin:{}'.format(vin))
            vin_code.type_keys(vin)
        # 申请单
        require_order = self._oper_window.child_window(auto_id="txtVBILLNOQ")
        require_order.type_keys(order)
        # 查询按钮
        query_btn = self._oper_window.child_window(title="查询(Q)", auto_id="button1")
        query_btn.click()
        self._logger.info('查询重大索赔审批单成功')
        # 结果-datagriview
        first_w = self._oper_window.child_window(title="行 0", control_type="Custom")
        # first_w.select()
        # first_w.double_click()
        # mouse.click()


class BaoYangManger(Handle):
    def __init__(self):
        super().__init__(work_name='保养管理')
        self.works = ['excel', 'query']

    def edit(self, query_para: BaoYangQuery, edit_para:BaoYangEdit):
        if len(query_para.date) == 2:
            start1, stop1 = query_para.date
            if start1 and stop1:
                start_date = self._oper_window.child_window(auto_id="dtBgeinDate")
                start_date.type_keys(start1)
                stop_date = self._oper_window.child_window(auto_id="dtEndDate")
                stop_date.type_keys(stop1)
            self._logger.debug('设置操作开始日期:{},结束日期:{}'.format(start1, stop1))
        else:
            self._logger.debug('操作日期为空，跳过设置')

        require_order = self._oper_window.child_window(auto_id="txtBillNo")

        if query_para.order:
            require_order.type_keys(query_para.order)
            self._logger.debug('设置订单号:{}'.format(query_para.order))
        query_btn = self._oper_window.child_window(title="查询(Q)", auto_id="button1")
        query_btn.click()
        self._logger.debug('执行查询完毕')
        # datagridview
        # dgv = self._oper_window.child_window(title="DataGridView", auto_id="dgvMaster", control_type="Table")
        dgv = self._oper_window.child_window(title="DataGridView", auto_id="dgvMaster", control_type="Table")

        # dgv_wrapper = dgv.wrapper_object()

        # mouse.double_click()

        # dgv_wrapper.double_click_input()
        # 移动到dgv的第一行，暂时没有dgv控件选中的功能
        mouse.double_click(button='left',coords=(451,267))
        edit_w = self._oper_window.child_window(title="维护区域", auto_id="Maintain", control_type="Group")
        edit_w.wait('ready',timeout=3)

        if edit_para.mark:
            edit_w.child_window(auto_id="txtVDLRREMARK").type_keys(edit_para.mark)
            self._logger.debug('设置备注:{}'.format(edit_para.mark))
        if edit_para.chepai:
            edit_w.child_window(auto_id="txtVLICENSE").type_keys(edit_para.chepai)
            self._logger.debug('设置车牌:{}'.format(edit_para.chepai))
        save_and_exit()
        if not self._oper_window.exists():
            self._logger.info("修改维护保养单成功，参数为：{}".format(edit_para))
        else:
            self._logger.info("可能UI卡顿了，开始执行尝试控件关闭".format(edit_para))
            try:
                self._oper_window.child_window(title="关闭", control_type="Button").click()
            except Exception as e:
                self._logger.exception("控件关闭发生异常:{}".format(e))
            else:
                self._logger.info("修改维护保养单成功，参数为：{}".format(edit_para))
        # 修改里程的时候可能有弹窗，esc取消
        # send_keys('{ESC}')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # h = SuoPeiService()
    # h.login()
    # 1.维护索赔单操作
    # h.init_window(node='维护索赔单')
    # h.query([], [], vin='LFNAHUMX2NAC09846', sel='撤回')
    # h.export_xlsx('suopei')
    # 2.维护重大索赔审批单操作
    # h.init_window(tree_item_title=h.work_name, node='维护重大索赔审批单')
    # h.query_zdspspt(oper_dates=[], order='ZD24060431923', vin='LFNAHULX8NAC21176', sel='返回')
    # h.export_xlsx('suopei')

    # 2.保养管理
    h = BaoYangManger()
    h.init_window(tree_item_title=h.work_name, node='维护保养单')
    para = BaoYangQuery(order='2212')
    edit_para = BaoYangEdit(mark='备注的测试')
    h.edit(query_para=para,edit_para=edit_para)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
