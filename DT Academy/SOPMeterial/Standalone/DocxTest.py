from docx import Document
import pandas as pd

OP_ROW = 1 
HEADER_ROW = 2 
FIRST_ROW = 3 

def getLeftHeaderCols(tb):
    op_cols = []
    ref_cols = []
    item_cols = []
    des_cols = []
    qty_cols = []
    for idx, _ in enumerate(tb.columns):
        if 'OP' in tb.cell(OP_ROW, idx).text:
            op_cols.append(idx)
        if 'Ref.' in tb.cell(HEADER_ROW, idx).text:
            ref_cols.append(idx)
        if 'Man.Item No.' in tb.cell(HEADER_ROW, idx).text:
            item_cols.append(idx)
        if 'Description' in tb.cell(HEADER_ROW, idx).text:
            des_cols.append(idx)
        if 'Qty.' in tb.cell(HEADER_ROW, idx).text:
            qty_cols.append(idx)
    if ref_cols:
        return {'ref':ref_cols[0],'item':item_cols[0],'des':des_cols[0],'qty':qty_cols[0]}

def getRightHeaderCols(tb):
    op_cols = []
    ref_cols = []
    item_cols = []
    des_cols = []
    qty_cols = []
    for idx, _ in enumerate(tb.columns):
        if 'OP' in tb.cell(OP_ROW, idx).text:
            op_cols.append(idx)
        if 'Ref.' in tb.cell(HEADER_ROW, idx).text:
            ref_cols.append(idx)
        if 'Man.Item No.' in tb.cell(HEADER_ROW, idx).text:
            item_cols.append(idx)
        if 'Description' in tb.cell(HEADER_ROW, idx).text:
            des_cols.append(idx)
        if 'Qty.' in tb.cell(HEADER_ROW, idx).text:
            qty_cols.append(idx)
            
    right_header_cols = {'ref': int ,'item':int,'des':int,'qty':int}
    try:
        right_header_cols['ref'] = ref_cols[2]
        for col in item_cols:
            if col > right_header_cols['ref']:
                right_header_cols['item'] = col
                break
        for col in des_cols:
            if col > right_header_cols['ref']:
                right_header_cols['des'] = col
                break
        for col in qty_cols:
            if col > right_header_cols['ref']:
                right_header_cols['qty'] = col
                break
    except:
        pass
    return right_header_cols

def GetLeftItemsList(table):
    item = []
    items = []
    headers = getLeftHeaderCols(table)
    for i,cells in enumerate(table.rows) : # left
        op = table.cell(OP_ROW, 3).text
        try:
            ref = table.cell(FIRST_ROW + i, headers['ref']).text
            item = table.cell(FIRST_ROW + i, headers['item']).text
            des = table.cell(FIRST_ROW  + i, headers['des']).text
            qty = table.cell(FIRST_ROW + i, headers['qty']).text
        except:
            continue
        if qty.isnumeric():
            item = [op, ref, item, des, qty]
            items.append(item)
    return items

def GetRightItemsList(table):
    item = []
    items = []
    headers = getRightHeaderCols(table)
    for i,cells in enumerate(table.rows) : # left
        op = table.cell(OP_ROW, 19).text
        try:
            ref = table.cell(FIRST_ROW + i, headers['ref']).text.replace('\n','')
            item = table.cell(FIRST_ROW + i, headers['item']).text
            des = table.cell(FIRST_ROW  + i, headers['des']).text
            qty = table.cell(FIRST_ROW + i, headers['qty']).text
        except:
            continue
        if qty.isnumeric():
            item = [op, ref, item, des, qty]
            items.append(item)
    return items


doc = Document('sop.docx')
df = pd.DataFrame(columns=['OperationStep','Ref.','Man.Item.No','Description','Qty'])

tb = doc.tables[71]
lefts = GetLeftItemsList(tb)
rights = GetRightItemsList(tb)

for left in lefts:
    df.loc[len(df)] = left
for right in rights:    
    df.loc[len(df)] = right

f_name = 'sop_docx.csv'
df.to_csv(f_name,encoding='utf-8-sig', index=False, mode='w', header=True)