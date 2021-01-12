import xlrd
import os
from xml.dom.minidom import Document

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

    def makeVariableXml(self):
        doc = Document() #创建DOM文档对象
        variableLibrary = doc.createElement('variable-library') #创建根元素
        doc.appendChild(variableLibrary)
        #传入变量
        category = doc.createElement('category')
        category.setAttribute('name', '传入变量')
        category.setAttribute('type', 'Custom')
        category.setAttribute('clazz', 'com.qsq.dmp.common.utils.DmpSyncMap')
        variableLibrary.appendChild(category)
        j = 1
        for i in range(self.rowNum-1):
            # 从第二行取对应 values 值
            values = self.table.row_values(j)
            #print(values)
            # 传入变量
            var = doc.createElement('var')
            var.setAttribute('act', 'InOut')
            var.setAttribute('source-type', values[0])
            var.setAttribute('label', values[1])
            var.setAttribute('name', values[2])
            var.setAttribute('type', values[3])
            category.appendChild(var)
            j += 1

        #临时变量
        category = doc.createElement('category')
        category.setAttribute('name', '临时变量')
        category.setAttribute('type', 'Custom')
        category.setAttribute('clazz', 'com.qsq.dmp.common.utils.DmpSyncMap')
        variableLibrary.appendChild(category)
        k = 1
        for i in range(self.temp_rowNum-1):
            # 从第二行取对应 values 值
            values = self.temp_table.row_values(k)
            #print(values)
            # 临时变量
            var = doc.createElement('var')
            var.setAttribute('act', 'InOut')
            var.setAttribute('source-type', values[0])
            var.setAttribute('label', values[1])
            var.setAttribute('name', values[2])
            var.setAttribute('type', values[3])
            category.appendChild(var)
            k += 1

        #写入xml文件
        f = open(self.variableXmlPath,'w',encoding='utf-8')
        #f.write(doc.toprettyxml(indent = '', newl='\n'))
        doc.writexml(f, addindent='\t', newl='\n',encoding='utf-8')
        f.close()

    def makeConstantXml(self):
        doc = Document()  # 创建DOM文档对象
        constantLibrary = doc.createElement('constant-library')  # 创建根元素
        doc.appendChild(constantLibrary)
        # 布尔
        category = doc.createElement('category')
        category.setAttribute('name', 'CL_BOOL_STATE')
        category.setAttribute('label', '布尔')
        constantLibrary.appendChild(category)
        constant = doc.createElement('constant')
        constant.setAttribute('name', 'true')
        constant.setAttribute('label', '真')
        constant.setAttribute('type', 'Boolean')
        category.appendChild(constant)
        constant = doc.createElement('constant')
        constant.setAttribute('name', 'false')
        constant.setAttribute('label', '假')
        constant.setAttribute('type', 'Boolean')
        category.appendChild(constant)
        # 执行结果
        category = doc.createElement('category')
        category.setAttribute('name', 'CL_EXEC_STATE')
        category.setAttribute('label', '执行结果')
        constantLibrary.appendChild(category)
        constant = doc.createElement('constant')
        constant.setAttribute('name', '0')
        constant.setAttribute('label', '通过')
        constant.setAttribute('type', 'Integer')
        category.appendChild(constant)
        constant = doc.createElement('constant')
        constant.setAttribute('name', '1')
        constant.setAttribute('label', '拒绝')
        constant.setAttribute('type', 'Integer')
        category.appendChild(constant)
        # 命中状态
        category = doc.createElement('category')
        category.setAttribute('name', 'CL_RULE_STATE')
        category.setAttribute('label', '命中状态')
        constantLibrary.appendChild(category)
        constant = doc.createElement('constant')
        constant.setAttribute('name', '0')
        constant.setAttribute('label', '未命中')
        constant.setAttribute('type', 'Integer')
        category.appendChild(constant)
        constant = doc.createElement('constant')
        constant.setAttribute('name', '1')
        constant.setAttribute('label', '命中')
        constant.setAttribute('type', 'Integer')
        category.appendChild(constant)
        #常量
        category = doc.createElement('category')
        category.setAttribute('name', 'CL_CONSTANT')
        category.setAttribute('label', '常量')
        constantLibrary.appendChild(category)
        j = 1
        for i in range(self.constant_rowNum - 1):
            values = self.constant_table.row_values(j)  # 从第二行取对应 values 值
            constant = doc.createElement('constant')
            constant.setAttribute('label', values[0])
            constant.setAttribute('name', values[1])
            constant.setAttribute('type', values[2])
            category.appendChild(constant)
            j += 1
        #写入xml文件
        f = open(self.constantXmlPath,'w',encoding='utf-8')
        doc.writexml(f, addindent='\t', newl='\n',encoding='utf-8')
        f.close()

    def makeParameterXmlPath(self):
        doc = Document()  # 创建DOM文档对象
        parameterLibrary = doc.createElement('parameter-library')  # 创建根元素
        doc.appendChild(parameterLibrary)
        parameter = doc.createElement('parameter')
        parameter.setAttribute('name', 'RuleResultSet')
        parameter.setAttribute('label', '规则结果集')
        parameter.setAttribute('type', 'List')
        parameter.setAttribute('act', 'InOut')
        parameterLibrary.appendChild(parameter)
        parameter = doc.createElement('parameter')
        parameter.setAttribute('name', 'FlowResultSet')
        parameter.setAttribute('label', '流程结果集')
        parameter.setAttribute('type', 'Map')
        parameter.setAttribute('act', 'InOut')
        parameterLibrary.appendChild(parameter)
        parameter = doc.createElement('parameter')
        parameter.setAttribute('name', 'ModelResultSet')
        parameter.setAttribute('label', '模型结果集')
        parameter.setAttribute('type', 'Map')
        parameter.setAttribute('act', 'InOut')
        parameterLibrary.appendChild(parameter)
        # 写入xml文件
        f = open(self.parameterXmlPath, 'w', encoding='utf-8')
        doc.writexml(f, addindent='\t', newl='\n', encoding='utf-8')
        f.close()

    def makeFeatureXmlPath(self):
        doc = Document()  # 创建DOM文档对象
        ruleSet = doc.createElement('rule-set')  # 创建根元素
        doc.appendChild(ruleSet)
        remark = doc.createElement('remark')
        remark_text = doc.createTextNode('<![CDATA[]]>')
        remark.appendChild(remark_text)
        ruleSet.appendChild(remark)
        rule = doc.createElement('rule')
        rule.setAttribute('name', 'QLZBLQ')
        ruleSet.appendChild(rule)
        remark = doc.createElement('remark')
        remark_text = doc.createTextNode('<![CDATA[ 全量指标拉取 ]]>')
        remark.appendChild(remark_text)
        rule.appendChild(remark)
        iif = doc.createElement('if')
        rule.appendChild(iif)
        aand = doc.createElement('and')
        iif.appendChild(aand)
        then = doc.createElement('then')
        rule.appendChild(then)
        eelse = doc.createElement('else')
        rule.appendChild(eelse)
        j = 1
        for i in range(self.feature_rowNum - 1):
            values = self.feature_table.row_values(j)  # 从第二行取对应 values 值
            varAssign = doc.createElement('var-assign')
            varAssign.setAttribute('var-category', '传入变量')
            varAssign.setAttribute('var', values[1])
            varAssign.setAttribute('var-label', values[0])
            varAssign.setAttribute('datatype', values[2])
            varAssign.setAttribute('type', 'variable')
            then.appendChild(varAssign)
            value = doc.createElement('value')
            value.setAttribute('bean-name', 'ApiBuiltInAction')
            value.setAttribute('bean-label', 'Api函数')
            value.setAttribute('method-name', 'apiIndexP2')
            value.setAttribute('method-label', '计算指标P2')
            value.setAttribute('type', 'Method')
            varAssign.appendChild(value)
            parameter = doc.createElement('parameter')
            parameter.setAttribute('name', '指标编码')
            parameter.setAttribute('type', 'String')
            value.appendChild(parameter)
            value1 = doc.createElement('value')
            value1.setAttribute('const-category', '常量')
            value1.setAttribute('const', values[1])
            value1.setAttribute('const-label', values[0])
            value1.setAttribute('type', 'Constant')
            parameter.appendChild(value1)
            parameter1 = doc.createElement('parameter')
            parameter1.setAttribute('name', '指标')
            parameter1.setAttribute('type', 'Object')
            value.appendChild(parameter1)
            value = doc.createElement('value')
            value.setAttribute('var-category', '传入变量')
            value.setAttribute('var', values[1])
            value.setAttribute('var-label', values[0])
            value.setAttribute('datatype', values[2])
            value.setAttribute('type', 'Variable')
            parameter1.appendChild(value)
            j += 1
        f = open(self.featureXmlPath, 'w', encoding='utf-8')
        doc.writexml(f, addindent='\t', newl='\n', encoding='utf-8')
        f.close()

if __name__ == "__main__":
    filePath = "./代码生成模板.xlsx"
    variableXmlPath = "./variable.xml"
    constantXmlPath = "./constant.xml"
    parameterXmlPath = "./parameter.xml"
    featureXmlPath = "./feature.xml"
    xmlMaker=XmlMaker(filePath,variableXmlPath,constantXmlPath,parameterXmlPath,featureXmlPath)
    xmlMaker.makeVariableXml()
    xmlMaker.makeConstantXml()
    xmlMaker.makeParameterXmlPath()
    xmlMaker.makeFeatureXmlPath()