#!/usr/bin/python
# -*- coding: UTF-8 -*-
import xlrd

class XmlMaker:
    def __init__(self, filePath, variableXmlPath, constantXmlPath, parameterXmlPath, featureXmlPath):
        self.filePath = filePath
        self.variableXmlPath = variableXmlPath
        self.constantXmlPath = constantXmlPath
        self.parameterXmlPath = parameterXmlPath
        self.featureXmlPath = featureXmlPath
        self.data = xlrd.open_workbook(self.filePath)  #打开 exlce 表格
        # table = data.sheets()[0] # 通过索引顺序获取
        # table = data.sheet_by_index(0) # 通过索引顺序获取

        self.table = self.data.sheet_by_name(u'传入变量')  # 通过名称获取
        self.rowNum = self.table.nrows  # 获取总行数
        #colNum = table.ncols  # 获取总列数

        self.temp_table = self.data.sheet_by_name(u'临时变量')  # 通过名称获取
        self.temp_rowNum = self.temp_table.nrows  # 获取总行数

        self.constant_table = self.data.sheet_by_name(u'常量')  # 通过名称获取
        self.constant_rowNum = self.constant_table.nrows  # 获取总行数

        self.feature_table = self.data.sheet_by_name(u'特征')  # 通过名称获取
        self.feature_rowNum = self.feature_table.nrows  # 获取总行数

    def makeXml(self):
        #生成传入变量
        f = open(self.variableXmlPath, 'w', encoding='utf-8')
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
        f.write("<variable-library>\n")
        f.write("<category name=\"传入变量\" type=\"Custom\" clazz=\"com.qsq.dmp.common.utils.DmpSyncMap\">\n")
        j=1
        for i in range(self.rowNum - 1):
            values = self.table.row_values(j)
            f.write("<var act=\"InOut\" name=\""+values[2]+"\" source-type=\""+values[0]+"\" label=\""+values[1]+"\" type=\""+values[3]+"\"/>\n")
            j += 1
        f.write("</category>\n")
        f.write("<category name=\"临时变量\" type=\"Custom\" clazz=\"com.qsq.dmp.common.utils.DmpSyncMap\">\n")
        j=1
        #生成临时变量
        for i in range(self.temp_rowNum - 1):
            values = self.temp_table.row_values(j)
            f.write("<var act=\"InOut\" name=\""+values[2]+"\" source-type=\""+values[0]+"\" label=\""+values[1]+"\" type=\""+values[3]+"\"/>\n")
            j += 1
        f.write("</category>\n")
        f.write("</variable-library>\n")
        f.close()

        #生成常量
        f = open(self.constantXmlPath, 'w', encoding='utf-8')
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
        f.write("<constant-library>\n")
        f.write("<category name=\"CL_BOOL_STATE\" label=\"布尔\">\n<constant name=\"true\" label=\"真\" type=\"Boolean\"/>\n")
        f.write("<constant name=\"false\" label=\"假\" type=\"Boolean\"/>\n</category>\n")
        f.write("<category name=\"CL_EXEC_STATE\" label=\"执行结果\">\n<constant name=\"0\" label=\"通过\" type=\"Integer\"/>\n")
        f.write("<constant name=\"1\" label=\"拒绝\" type=\"Integer\"/>\n</category>")
        f.write("<category name=\"CL_RULE_STATE\" label=\"命中状态\">\n<constant name=\"0\" label=\"未命中\" type=\"Integer\"/>\n")
        f.write("<constant name=\"1\" label=\"命中\" type=\"Integer\"/>\n</category>\n")
        f.write("<category name=\"CL_CONSTANT\" label=\"常量\">\n")
        j=1
        for i in range(self.constant_rowNum - 1):
            values = self.constant_table.row_values(j)
            f.write("<constant label=\""+values[0]+"\" name=\""+values[1]+"\" type=\""+values[2]+"\"/>\n")
            j += 1
        f.write("</category>\n")
        f.write("</constant-library>\n")
        f.close()

        #生成参数
        f = open(self.parameterXmlPath, 'w', encoding='utf-8')
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
        f.write("<parameter-library>\n")
        f.write("<parameter name=\"RuleResultSet\" label=\"规则结果集\" type=\"List\" act=\"InOut\"/>\n")
        f.write("<parameter name=\"FlowResultSet\" label=\"流程结果集\" type=\"Map\" act=\"InOut\"/>\n")
        f.write("<parameter name=\"ModelResultSet\" label=\"模型结果集\" type=\"Map\" act=\"InOut\"/>\n")
        f.write("</parameter-library>\n")
        f.close()

        #指标拉取
        f = open(self.featureXmlPath, 'w', encoding='utf-8')
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
        f.write("<rule-set>\n<remark><![CDATA[]]></remark>\n<rule name=\"QLZBLQ\">\n<remark><![CDATA[ 全量指标拉取 ]]></remark>\n<if>\n<and/>\n</if>\n<then>\n")
        j=1
        for i in range(self.feature_rowNum - 1):
            values = self.feature_table.row_values(j)
            f.write("<var-assign var-category=\"传入变量\" var=\""+values[1]+
                    "\" var-label=\""+values[0]+"\" datatype=\""+values[2]+"\" type=\"variable\">\n")
            f.write("<value bean-name=\"ApiBuiltInAction\" bean-label=\"Api函数\" method-name=\"apiIndexP2\""
                    " method-label=\"计算指标P2\" type=\"Method\">\n<parameter name=\"指标编码\" type=\"String\">\n")
            f.write("<value const-category=\"常量\" const=\""+values[1]+"\" const-label=\""+values[0]+"\" type=\"Constant\"/>\n")
            f.write("</parameter>\n<parameter name=\"指标\" type=\"Object\">\n")
            f.write("<value var-category=\"传入变量\" var=\""+values[1]+
                    "\" var-label=\""+values[0]+"\" datatype=\""+values[2]+"\" type=\"Variable\"/>\n")
            f.write("</parameter>\n</value>\n</var-assign>\n")
            j += 1
        f.write("</then>\n<else/>\n</rule>\n</rule-set>\n")
        f.close()

if __name__ == "__main__":
    filePath = "./代码生成模板.xlsx"
    variableXmlPath = "./variable.xml"
    constantXmlPath = "./constant.xml"
    parameterXmlPath = "./parameter.xml"
    featureXmlPath = "./feature.xml"
    xmlMaker=XmlMaker(filePath,variableXmlPath,constantXmlPath,parameterXmlPath,featureXmlPath)
    xmlMaker.makeXml()