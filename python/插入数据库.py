import re,json,pyperclip,time,random,cv2,pypyodbc
def 产品库导入(biao='SelfParamUnit',path= r'C:\Users\Admin\Documents\测试客餐厅\场景文件\Import.mdb'):

	lie = shujuku(sql = f'select * from {biao}',path=path,q = 3)
	
	s1 = pyperclip.paste().split("\r\n")[:-1]
	for x in s1:
		lie1 =lie.copy()
		charushuju = x.split('\t')[1:]
		# print(charushuju)
		for y in range(0,len(charushuju)):
			
			if charushuju[y] =='' or charushuju[y] ==None :
				# print(charushuju[y])
				# print(y)
				lie1.remove(lie[y])
				# print(lie1)
		# print(len(lie1))
		charushuju = list( filter(lambda a: a!='' and a!=None ,charushuju) )
		lieshu = len(charushuju)
		lie1 = lie1[:lieshu]
		sql = f"INSERT INTO {biao} ({','.join(lie1)}) VALUES ("
		sql = sql+str(charushuju)[1:-1].replace('\\\\','\\')+')'
		print(sql)
		shujuku(sql =sql,path = path )
def shujuku(path,sql,q = 1):
	# path=r'E:\er\py\v20查数据\V20产品库\产品库.accdb'# 数据库文件
	t = []
	conn = pypyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + path + ";Uid=;Pwd=;")
	cursor = conn.cursor()
	# sql = f"INSERT INTO chanpingku1 (id, path_id, dir1, dir2, dir3, dir4, modelno, name, l, w, h, product_type, param_type, paramunit_name, model_path, model_id, model168, pic, change_size, all_doorstyle, reverse_permit_door_mode, door_style_scheme_id, direction, brand, worktop_mx_line, btm_light_line, elva_mode, elva, elva_cfg, need_quotation, code, standard, color, unit, price, material, comment, class, ext1, net_id, vr_lib_extra_info, version, hide, status, spzp, wayes, homkoo, created_at, updated_at, deleted_at, updated_by, params) VALUES ("
	# print(sql+str(t)[1:-1]+',)')
	# print(sql)
	SQL_value = 'select * from chanpingku1'#
	# SQL_value = 'delete * from chanpingku1'

	if q==2:
		cursor.execute(sql)
		# for row in cursor.execute(sql):
		#     print(row)
		t =  list(cursor.fetchall()   )
	elif q == 3:
		cursor.execute(sql)
		# for row in cursor.description[1:]:
		    # print(row[0])    
		t =list(map(lambda a:a[0],cursor.description[1:]))
	else:
		cursor.execute(sql)
		t = 1
	# cursor.execute(SQL_value)
	cursor.commit()
	conn.close()
	return t


产品库导入()