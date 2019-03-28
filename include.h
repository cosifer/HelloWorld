#pragma once
#include <QList>
#include <QTableWidget>

//项目信息 添加注释（如何解决冲突）

struct ProjInfo
{
	char ID[50];		//项目索引
	char Name[100];		//项目名称
	//char state[20];		//立项情况
	int	 iYear;			//年份
	char PC[20];		//批次
	char type[20];		//资金性质
	char CoLtd[50];		//施工单位
	double fTZ;			//项目总投资
	char Time[20];		//项目创建时间
};



static void MapDataToTable(const QStringList arrField, const QList<QStringList>& arrData, QTableWidget* pTable, const QStringList arrInt, const QStringList arrFloat)
{
	if (pTable && arrField.size() > 0)
	{
		//pTable->setRowCount(0);
		//pTable->setColumnCount(0);

		int iMaxRowNum = arrData[0].size();
		for(int i = 0; i < arrData.size(); i++)
		{
			if(arrData[i].size() > iMaxRowNum)
			{
				iMaxRowNum = arrData[i].size();
			}
		}
		pTable->setRowCount(iMaxRowNum);
		pTable->setColumnCount(arrField.size());

		for (int i = 0; i < arrField.size(); i++)
		{
			//添加表头
			QTableWidgetItem* pHeaderItem = new QTableWidgetItem(arrField[i]);
			pHeaderItem->setTextAlignment(Qt::AlignCenter);
			pTable ->setHorizontalHeaderItem(i, pHeaderItem);

			//添加列数据
			QStringList arrValue = arrData[i];
			for (int j=0; j < arrValue.size(); j++)
			{
				QTableWidgetItem* pItem = new QTableWidgetItem(arrValue[j]);
				pTable->setItem(j,i, pItem);
			}
		}
	}
}

static void MapDataToValueTable(const QStringList arrField, const QMap<QString,bool> arrFielvVisable, const QList<QStringList>& arrData, QTableWidget* pTable, const QStringList arrInt, const QStringList arrFloat, QList<double>& arrZJ)
{
	if (pTable && arrField.size() > 0)
	{
		pTable->setRowCount(0);
		pTable->setColumnCount(0);

		int iMaxRowNum = 0;
		for(int i = 0; i < arrData.size(); i++)
		{
			if(arrData[i].size() > iMaxRowNum)
			{
				iMaxRowNum = arrData[i].size();
			}
		}
		pTable->setRowCount(iMaxRowNum+3);
		pTable->setColumnCount(arrField.size());

		//设置表格表头
		for(int j = 0; j < 3; j++)
		{
			for(int i = 0; i < arrField.size(); i++)
			{
				QTableWidgetItem* pItem = new QTableWidgetItem(QString::fromLocal8Bit(""));
				pItem->setTextAlignment(Qt::AlignCenter);
				QFont fnt = pItem->font();
				fnt.setBold(true);
				pItem->setFont(fnt);
				pTable->setItem(j,i,pItem);
			}
		}
		pTable->setSpan(0,0,3,1);
		pTable->setSpan(0,1,3,1);
		pTable->setSpan(0,2,3,1);
		pTable->setSpan(0,3,3,1);
		pTable->setSpan(0,4,3,1);
		pTable->setSpan(0,5,3,1);
		pTable->setSpan(0,6,3,1);
		pTable->item(0,0)->setText(QString::fromLocal8Bit("物资编码&扩展码"));
		pTable->item(0,1)->setText(QString::fromLocal8Bit("大分类"));
		pTable->item(0,2)->setText(QString::fromLocal8Bit("中分类"));
		pTable->item(0,3)->setText(QString::fromLocal8Bit("小分类"));
		pTable->item(0,4)->setText(QString::fromLocal8Bit("物料描述"));
		pTable->item(0,5)->setText(QString::fromLocal8Bit("扩展描述"));
		pTable->item(0,6)->setText(QString::fromLocal8Bit("技术规范ID"));

		//物资协议表
		pTable->setSpan(0,7,1,2);
		pTable->setSpan(1,7,1,2);
		pTable->item(0,7)->setText(QString::fromLocal8Bit("协议物资表"));
		pTable->item(2,7)->setText(QString::fromLocal8Bit("单位"));
		pTable->item(2,8)->setText(QString::fromLocal8Bit("单价"));
		
		//物资清册
		pTable->setSpan(0,9,1,3);
		pTable->setSpan(1,9,1,3);
		pTable->item(0,9)->setText(QString::fromLocal8Bit("物资清册"));
		pTable->item(2,9)->setText(QString::fromLocal8Bit("单位"));
		pTable->item(2,10)->setText(QString::fromLocal8Bit("数量"));
		pTable->item(2,11)->setText(QString::fromLocal8Bit("单价"));

		//施工清册
		pTable->setSpan(0,12,1,1);
		pTable->setSpan(1,12,1,1);
		pTable->item(0,12)->setText(QString::fromLocal8Bit("施工清册"));
		pTable->item(2,12)->setText(QString::fromLocal8Bit("数量"));

		//结算清册
		pTable->setSpan(0,13,1,2);
		pTable->setSpan(1,13,1,2);
		pTable->item(0,13)->setText(QString::fromLocal8Bit("结算清册"));
		pTable->item(2,13)->setText(QString::fromLocal8Bit("单位"));
		pTable->item(2,14)->setText(QString::fromLocal8Bit("数量"));

		//ERP
		pTable->setSpan(0,15,1,3);
		pTable->setSpan(1,15,1,3);
		pTable->item(0,15)->setText("ERP");
		pTable->item(2,15)->setText("单位");
		pTable->item(2,16)->setText("单价");
		pTable->item(2,17)->setText("数量");

		for (int i = 0; i < arrField.size(); i++)
		{
			//添加列数据
			QStringList arrValue = arrData[i];
			for (int j=0; j < arrValue.size(); j++)
			{
				QTableWidgetItem* pItem = new QTableWidgetItem(arrValue[j]);
				pTable->setItem(j+3,i, pItem);
			}
			
			if(!arrFielvVisable[arrField.at(i)])
			{
				pTable->hideColumn(i);
			}
		}

		/*
		//统计各清册总价
		double fWZBZJ= 0;	//物资表总价
		double fWZQCZJ= 0; //物资清册总价
		double fSGZJ= 0;	//施工清册总价
		double fJSZJ= 0;	//结算清册总价
		double fERPZJ=0;	//ERP总价
		int iRow = arrData[0].size();
		for (int i = 0; i < iRow; i++)
		{
			//添加列数据
			for (int j=0; j < arrField.size(); j++)
			{
				double fWZBDJ = arrData[8] .at(j).toDouble();	//物资表单价
				double fQCDJ  = arrData[11].at(j).toDouble();	//物资清册单价
				double fQCSL  = arrData[10].at(j).toDouble();	//物资清册数量
				double fSGSL  = arrData[12].at(j).toDouble();	//施工清册数量
				double fJSSL  = arrData[14].at(j).toDouble();	//结算清册数量
				double fERPDJ = arrData[16].at(j).toDouble();	//ERP单价
				double fERPSL = arrData[17].at(j).toDouble();	//ERP数量

				fWZBZJ		+= fWZBDJ*fQCSL;
				fWZQCZJ		+= fQCDJ*fQCSL;
				fSGZJ		+= fQCDJ*fQCSL;
				fJSZJ		+= fQCDJ*fQCSL;
				fERPZJ		+= fERPDJ*fERPSL;
			}
		}

		arrZJ << fWZBZJ <<fWZQCZJ << fSGZJ << fJSZJ << fERPZJ;
		*/

		pTable->setItem(1,7,	new QTableWidgetItem("总价: "+QString::number(arrZJ[0],'f',2)));
		pTable->setItem(1,9,	new QTableWidgetItem("总价: "+QString::number(arrZJ[1],'f',2)));
		pTable->setItem(1,12,	new QTableWidgetItem("总价: "+QString::number(arrZJ[2],'f',2)));
		pTable->setItem(1,13,	new QTableWidgetItem("总价: "+QString::number(arrZJ[3],'f',2)));
		//pTable->setItem(1,12, new QTableWidgetItem("总价"));
		//pTable->setItem(1,13, new QTableWidgetItem("总价"));
		pTable->setItem(1,15,	new QTableWidgetItem("总价: "+QString::number(arrZJ[4],'f',2)));
	}
}

static void TableData2Map(QTableWidget* pTable, QStringList& arrOldField, QList<QStringList>& arrOldData, QStringList& arrNewField, QList<QStringList>& arrNewData)
{
	int iRowCnt = pTable->rowCount();
	for(int i = 0; i < arrOldField.size(); i++)
	{
		arrNewField << pTable->horizontalHeaderItem(i)->text();

		QString strData;
		for(int j = 0; j < iRowCnt; j++)
		{
			strData += pTable->item(j,i)->text();
			if(j != iRowCnt-1)
			{
				strData += QString::fromLocal8Bit(";");
			}
		}
		QStringList arrData = strData.split(";");
		arrNewData << arrData;
	}
}

static void ValueTableData2Map(QTableWidget* pTable, QStringList& arrOldField, QList<QStringList>& arrOldData, QStringList& arrNewField, QList<QStringList>& arrNewData)
{
	int iRowCnt = pTable->rowCount();
	for(int i = 0; i < arrOldField.size(); i++)
	{
		QString strData;
		for(int j = 2; j < iRowCnt; j++)
		{
			strData += pTable->item(j,i)->text();
			if(j != iRowCnt-1)
			{
				strData += QString::fromLocal8Bit(";");
			}
		}
		QStringList arrData = strData.split(";");
		arrNewData << arrData;
	}
}

#include <QMessageBox>
//iType 0项目、1物资、2施工、3结算、4物资表
static bool check(QString iType, QStringList arrField)
{
	QStringList arr;
	if (iType == "@XM")
	{
		arr << QString::fromLocal8Bit("年份") << QString::fromLocal8Bit("批次") << QString::fromLocal8Bit("设计索引号")<<QString::fromLocal8Bit("设计单位")<<QString::fromLocal8Bit("项目名称")<<QString::fromLocal8Bit("资金属性")
			<< QString::fromLocal8Bit("项目总投资");
	}
	else if (iType == "@WZ")
	{
		arr <<"类别编码" << "物资类别"<< "物资编码" << "扩展码" <<"物料描述"<<"扩展码描述"<<"物资性质"<<"设计单位"<<"设计数量"
			<<"上报单位"<<"上报数量"<<"单价"<<"合价（元）"<<"单重\n（T）"<<"合重\n（T）";
	}
	else if(iType == "@SG")
	{
		arr << QString::fromLocal8Bit("物资编码") << QString::fromLocal8Bit("扩展码")<<QString::fromLocal8Bit("施工上报数量")<<QString::fromLocal8Bit("施工上报单价");
	}
	else if (iType == "@JS")
	{
		arr << QString::fromLocal8Bit("物资编码") << QString::fromLocal8Bit("扩展码")<<QString::fromLocal8Bit("结算数量")<<QString::fromLocal8Bit("结算总价");
	}
	else if(iType == "@WZB")
	{
		arr << QString::fromLocal8Bit("物料&扩展") << QString::fromLocal8Bit("大分类")<<QString::fromLocal8Bit("中分类")<<QString::fromLocal8Bit("小分类")
			<<QString::fromLocal8Bit("技术规范ID")<<QString::fromLocal8Bit("单位")<<QString::fromLocal8Bit("价格");
	}

	bool bRet = true;
	QString strMsg = QString::fromLocal8Bit("缺少");
	for (int i = 0; i<arr.size(); i++)
	{
		if(!arrField.contains(arr[i]))
		{
			bRet = false;
			strMsg += arr[i];
			strMsg += QString::fromLocal8Bit("、");
		}
	}
	if(!bRet)
	{
		strMsg = strMsg.mid(0,strMsg.size()-1);
		strMsg += QString::fromLocal8Bit("字段");
		QMessageBox::information(0, QString::fromLocal8Bit("提示"), strMsg,QMessageBox::Ok);
	}

	return bRet;
}