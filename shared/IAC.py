"""
(Purpose) IAC.py is a module that contains functions used in the IAC report
"""


import math
from lxml import etree
import latex2mathml.converter

def add_thousand_sep(dic):
    """
    Add thousand separator to numbers in a dictionary and format it to string
    :param dic: Dictionary
    :return: Dictionary with keys in thousand separator
    """
    for key in dic.keys():
        if type(dic[key]) == int or type(dic[key]) == float:
            dic[key] = f'{dic[key]:,}'
        else:
            pass
    return dic


# Add equations
def add_eqn(doc, eqn, eqn_input):
    """
    Add equation to Word document, search for eqn in doc and replace with eqn_input
    :param doc: Document
    :param eqn: Equation string
    :param eqn_input: Equation input
    :return: None
    """
    for p in doc.paragraphs:
        if eqn in p.text:
            p.text = p.text.replace(eqn, '')
            word_math = latex_to_word(eqn_input)
            p._element.append(word_math)


# Function to create Word equation from LaTeX
def latex_to_word(latex_input):
    """
    Convert LaTeX equation to Word equation
    :param latex_input: LaTeX equation
    :return: Word equation
    """
    mathml = latex2mathml.converter.convert(latex_input)
    tree = etree.fromstring(mathml)
    xslt = etree.parse('../shared/MML2OMML.XSL')
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    return new_dom.getroot()

# Payback Period formatting
def payback(ACS,IC):
    """
    Formet payback period by year and month
    :param ACS: Annual Cost Savings ($/yr)
    :param IC: Implementation Cost ($)
    :return: formatted Payback Period (str)
    """
    PB = IC / ACS
    if PB <= 11.0 / 12.0:
        PB = math.ceil(PB * 12.0)
        PBstr = str(PB) + " month"
    else:
        PB = math.ceil(PB * 10.0) / 10.0
        PBstr = str(PB).rstrip("0").rstrip(".") + " year"
    if PB > 1.0:
        PBstr = PBstr + "s"
    return PBstr