from pycel import ExcelCompiler
import pycel
import logging

filename = r'carlin_soskice_macroeconomic_simulator (unprotected) (2).xlsx'


# load & compile the file to a graph
excel = ExcelCompiler(filename=filename,cycles=True)

eval_ctx = pycel.excelformula.ExcelFormula.build_eval_context(
                excel._evaluate, excel._evaluate_range,
                excel.log, plugins=excel._plugin_modules)

def eval_formula(formula): 
    return eval_ctx(pycel.excelformula.ExcelFormula(formula), cse_array_address=None) 

print(excel.evaluate('main page!G11'))
print(excel.evaluate('normal case!L3:L30'))

excel.set_value('main page!G11', 8)
print(excel.evaluate('main page!G11'))
print(excel.evaluate('normal case!L3:L30', iterations=100, tolerance=0.001))

#excel.recalculate()
print(excel.evaluate('normal case!L3', iterations=100, tolerance=0.001))
print(excel.evaluate('normal case!L4'))
print(excel.evaluate('normal case!L5'))
print(excel.evaluate('normal case!L6'))
print(excel.evaluate('normal case!L7'))

print(excel.evaluate('normal case!L3:L31', iterations=100, tolerance=0.001))
