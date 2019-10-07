var dataFile = "/medias/assets/data/data.xlsx";
var oReq;
if (window.XMLHttpRequest) oReq = new XMLHttpRequest();
else if (window.ActiveXObject) oReq = new ActiveXObject('MSXML2.XMLHTTP.3.0');
else throw "XHR unavailable for your browser";
oReq.open("GET", dataFile, true);

if (typeof Uint8Array !== 'undefined') {
	oReq.responseType = "arraybuffer";
	oReq.onload = function (e) {
		var arraybuffer = oReq.response;
		var data = new Uint8Array(arraybuffer);
		var dataRaw = XLSX.read(data, {
			type: "array"
		});
		dataJson = toJson(dataRaw);
		drawTables();
		drawCharts();
	};
}
oReq.send();

function toJson(workbook) {
	var result = [];
	workbook.SheetNames.forEach(function (sheetName) {
		var json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
		if (json.length > 0) {
			result.push(json);
		}
	});
	return result;
};

function drawCharts() {
	function arrSum(arr) {
		return arr.reduce(function (prev, curr, idx, arr) {
			return Math.round((prev + curr) * 100) / 100;
		});
	}
	var couLabels = [];
	var arrIn = [];
	var arrOut = [];
	var arrDiff = [];
	var alipayIn = [];
	var wechatIn = [];
	var bilibiliIn = [];
	var paypalIn = [];
	var otherIn = [];
	for (var i in dataJson[2]) {
		couLabels.push.apply(couLabels, [dataJson[2][i].MONTH + '月']);
		arrIn.push.apply(arrIn, [dataJson[2][i].IN.toFixed(2)]);
		arrOut.push.apply(arrOut, [-dataJson[2][i].OUT.toFixed(2)]);
		arrDiff.push.apply(arrDiff, [dataJson[2][i].DIFF.toFixed(2)]);
		alipayIn.push.apply(alipayIn, [dataJson[2][i].ALIPAY]);
		wechatIn.push.apply(wechatIn, [dataJson[2][i].WECHAT]);
		bilibiliIn.push.apply(bilibiliIn, [dataJson[2][i].BILIBILI]);
		paypalIn.push.apply(paypalIn, [dataJson[2][i].PAYPAL]);
		otherIn.push.apply(otherIn, [dataJson[2][i].OTHER]);
	}
	var ansIn = new Array()
	ansIn[0] = arrSum(alipayIn)
	ansIn[1] = arrSum(wechatIn)
	ansIn[2] = arrSum(bilibiliIn)
	ansIn[3] = arrSum(paypalIn)
	ansIn[4] = arrSum(otherIn)
	var mixChartData = {
		chart: {
			height: 350,
			type: 'line',
			toolbar: {
				show: false
			}
		},
		series: [{
			name: '收入',
			type: 'column',
			data: arrIn
		}, {
			name: '支出',
			type: 'column',
			data: arrOut
		}, {
			name: '结余',
			type: 'line',
			data: arrDiff
		}],
		stroke: {
			width: [0, 0, 5],
			curve: 'smooth'
		},
		title: {
			text: '统计',
			style: {
				fontSize: '18px'
			}
		},
		labels: couLabels,
		xaxis: {
			labels: {
				hideOverlappingLabels: true
			}
		}
	}
	var mixChartData = new ApexCharts(
		document.querySelector("#mixChart"),
		mixChartData
	);
	mixChartData.render();


	var inAnsData = {
		chart: {
			type: 'donut',
			width: '90%'
		},
		dataLabels: {
			enabled: true,
		},
		plotOptions: {
			pie: {
				donut: {
					size: '70%',
				},
				offsetY: 30,
			},
			stroke: {
				colors: undefined
			}
		},
		colors: ['#00D8B6', '#008FFB', '#FEB019', '#FF4560', '#775DD0'],
		title: {
			text: '收入分析',
			style: {
				fontSize: '18px'
			}
		},
		series: ansIn,
		labels: ["支付宝", "微信", "哔哩哔哩", "PayPal", "其他"],
		legend: {
			position: 'top',
			offsetY: 10
		}
	}
	var inAnsData = new ApexCharts(
		document.querySelector("#inAns"),
		inAnsData
	)
	inAnsData.render();

};

function drawTables() {
	function arrSum(arr) {
		return arr.reduce(function (prev, curr, idx, arr) {
			return Math.round((prev + curr) * 100) / 100;
		});
	}
	tableIn = null;
	tableOut = null;
	for (var i in dataJson[0]) {
		var date = dataJson[0][i].DATE.toString()
		var date = date.slice(0, 4) + '年' + date.slice(4, -2) + '月' + date.slice(-2) + '日';
		tableIn += '<tr><td>' + date + '</td><td>' + dataJson[0][i].USER + '</td><td>' + dataJson[0][i].PLATFORM + '</td><td>￥' + dataJson[0][i].MONEY + '</td></tr>';
		var date = null;
	}
	for (var i in dataJson[1]) {
		var date = dataJson[1][i].DATE.toString()
		var date = date.slice(0, 4) + '年' + date.slice(4, -2) + '月' + date.slice(-2) + '日';
		tableOut += '<tr><td>' + date + '</td><td>' + dataJson[1][i].ADMIN + '</td><td>' + dataJson[1][i].ITEM + '</td><td>￥' + dataJson[1][i].MONEY + '</td></tr>';
	}
	$('#inTable tbody').append(tableIn);
	$('#outTable tbody').append(tableOut);
	$('#inTable, #outTable').DataTable({
		bFilter: false,
		bLengthChange: false,
		"orderFixed":[
			[0, "desc"],
			[1, "desc"],
			[2, "desc"],
			[3, "desc"]
		],
		language: {
			url: '/medias/assets/lang/zh-cn.json'
		}
	});
	valIn = [];
	valOut = [];
	valDiff = [];
	for (var i in dataJson[2]) {

		valIn.push.apply(valIn, [dataJson[2][i].IN]);
		valOut.push.apply(valOut, [dataJson[2][i].OUT]);
		valDiff.push.apply(valDiff, [dataJson[2][i].DIFF]);

	}
	$('#sumIn').html("总收入 : ￥" + arrSum(valIn));
	$('#sumOut').html("总支出 : ￥" + arrSum(valOut));
	$('#sumDiff').html("总结余 : ￥" + arrSum(valDiff));
}