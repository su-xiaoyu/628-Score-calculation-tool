function selectFile() {
	document.getElementById('file').click()
}

// 读取本地excel文件
function readWorkbookFromLocalFile(file, callback) {
	var reader = new FileReader()
	reader.onload = function (e) {
		var data = e.target.result
		var workbook = XLSX.read(data, {
			type: 'binary',
		})
		if (callback) callback(workbook)
	}
	reader.readAsBinaryString(file)
}
// 读取 excel文件
function outputWorkbook(workbook) {
	var sheetNames = workbook.SheetNames // 工作表名称集合
	sheetNames.forEach((name) => {
		var worksheet = workbook.Sheets[name] // 只能通过工作表名称来获取指定工作表
		for (var key in worksheet) {
			// v是读取单元格的原始值
			console.log(key, key[0] === '!' ? worksheet[key] : worksheet[key].v)
		}
	})
}

function readWorkbook(workbook) {
	var sheetNames = workbook.SheetNames // 工作表名称集合
	var worksheet = workbook.Sheets[sheetNames[0]] // 这里我们只读取第一张sheet
	//var csv = XLSX.utils.sheet_to_csv(worksheet);
	//document.getElementById('result').innerHTML = csv2table(csv);
	var myjson_1 = XLSX.utils.sheet_to_json(worksheet),
		myjson_2 = XLSX.utils.sheet_to_json(worksheet),
		myjson_3 = XLSX.utils.sheet_to_json(worksheet)

	for (let i = 0; i < myjson_1.length; i++) {
		myjson_1[i].change = -(myjson_1[i].last - myjson_1[i].current)
		myjson_2[i].change = -(myjson_2[i].last - myjson_2[i].current)
		myjson_3[i].change = -(myjson_3[i].last - myjson_3[i].current)
		//console.log(myjson[i])
	}
	//console.log(myjson[5].change)

	/*获取指定单元格*/
	var cell_1 = 'A1',
		cell_2 = 'B1',
		cell_3 = 'C1'
	/* Find desired cell */
	var desired_cell_1 = worksheet[cell_1],
		desired_cell_2 = worksheet[cell_2],
		desired_cell_3 = worksheet[cell_3]

	/* Get the value */
	var value_1 = desired_cell_1 ? desired_cell_1.v : undefined,
		value_2 = desired_cell_2 ? desired_cell_2.v : undefined,
		value_3 = desired_cell_3 ? desired_cell_3.v : undefined
	//var name = value_1
	//console.log(typeof (name)) //姓名name
	//console.log(value_2) //这一次current
	//console.log(value_3) //上一次last

	/* var obj = Object.assign(myjson1)
    console.log(obj) */
	//console.log(myjson1)

	function compare1(value) {
		return function (a, b) {
			var a = a[value]
			var b = b[value]
			return b - a
		}
	} //降序

	function compare2(value) {
		return function (a, b) {
			var a = a[value]
			var b = b[value]
			return a - b
		}
	} //升序
	var obj1 = myjson_1.sort(compare1('change')),
		obj2 = myjson_2.sort(compare2('current')),
		obj3 = myjson_3.sort(compare1('last'))
	console.log(obj1)
	console.log(obj2)

	var array = new Array()
	for (let i = 0, j = 0; i < myjson_1.length; i++) {
		//console.log(obj2[i])
		array[j] = obj2[i]
		j += 1
		//console.log(obj1[i])
		array[j] = obj1[i]
		j += 1
	}
	//console.log(array)

	function remove(arr) {
		for (var i = 0; i < arr.length; i++) {
			for (var j = i + 1; j < arr.length; j++) {
				if (arr[i].name == arr[j].name) {
					//第一个等同于第二个，splice方法删除第二个
					arr.splice(j, 1)
					j--
				}
			}
		}
		return arr
	} //查重
	var result_arr = remove(array)
	//console.log(result_arr)

	var group_1 = new Array(),
		group_2 = new Array(),
		group_3 = new Array(),
		group_4 = new Array(),
		group_5 = new Array(),
		group_6 = new Array()

	for (let i = 0; i < result_arr.length; i++) {
		if (i <= 9) {
			group_1[i] = obj3[i]
		} else if (i > 9 && i <= 19) {
			group_2[i - 10] = obj3[i]
		} else if (i > 19 && i <= 29) {
			group_3[i - 20] = obj3[i]
		} else if (i > 29 && i <= 39) {
			group_4[i - 30] = obj3[i]
		} else if (i > 39 && i <= 49) {
			group_5[i - 40] = obj3[i]
		} else {
			group_6[i - 50] = obj3[i]
		}
	}
	var group_1 = group_1.sort(compare1('change')),
		group_2 = group_2.sort(compare1('change')),
		group_3 = group_3.sort(compare1('change')),
		group_4 = group_4.sort(compare1('change')),
		group_5 = group_5.sort(compare1('change')),
		group_6 = group_6.sort(compare1('change'))

	console.log(group_1)
	console.log(group_2)
	console.log(group_3)
	console.log(group_4)
	console.log(group_5)
	console.log(group_6)
	group_6[9] = {} //解决下面的二维数组[5][9]未定义

	var group = [group_1, group_2, group_3, group_4, group_5, group_6]

	console.log(group)
	var n = 1
	var result_2 = document.getElementById('result_2')
	for (let j = 0; j < group[1].length; j++) {
		for (let i = 0; i < group.length; i++) {
			if (j == 9 && i == 5) {
				break
			}
			//console.log(i+"+"+j)

			var txt = document.createTextNode(
				n + '. ' + group[i][j].name + '    变化：' + group[i][j].change
			)
			var p = document.createElement('p')
			p.appendChild(txt)
			result_2.appendChild(p)
			n += 1
		}
	}

	var result_1 = document.getElementById('result_1')
	for (let i = 0; i < group.length; i++) {
		var n = 1
		for (let j = 0; j < group[1].length; j++) {
			if (j == 9 && i == 5) {
				break
			}
			//console.log(i+";;"+j)

			//console.log(". " + group[i][j].name + "    变化：" + group[i][j].change)

			var txt1 = document.createTextNode(
				n + '. ' + group[i][j].name + '    变化：' + group[i][j].change
			)
			var p = document.createElement('p')
			p.appendChild(txt1)
			result_1.appendChild(p)
			n += 1
			if ((j + 1) % 10 == 0) {
				var txt = document.createTextNode('------------------')
				var p = document.createElement('p')
				p.appendChild(txt)
				result_1.appendChild(p)
			}
		}
	}
}

$(function () {
	document.getElementById('file').addEventListener('change', function (e) {
		var files = e.target.files
		if (files.length == 0) return
		var f = files[0]
		if (!/\.xlsx$/g.test(f.name)) {
			alert('仅支持读取xlsx格式！')
			return
		}
		readWorkbookFromLocalFile(f, function (workbook) {
			readWorkbook(workbook)
		})
	})
})
