#include "ExcelProcess.h"

#include <QMessageBox>
#include <QTableWidget>
#include <QTableWidgetItem>

ExcelProcess::ExcelProcess(QObject *parent)
	: QObject(parent)
{
	//需要读取字段汇总
	m_WZFlags <<"类别编码" << "物资类别"<< "物资编码" << "扩展码" <<"物料描述"<<"扩展码描述"<<"物资性质"<<"设计单位"<<"设计数量"
		<<"上报单位"<<"上报数量"<<"单价"<<"合价（元）"<<"单重\n（T）"<<"合重\n（T）";

	m_WZBFlags << "物料&扩展" << "大分类" << "中分类" << "小分类" << "物料编码" << "物料描述" << "扩展编码" << "扩展描述" << "单位"
		<< "技术规范ID" << "价格" << "最终确认标段";

	//需要进行异常字符判断的字符
	m_XMCheck << "项目设计索引号" << "项目名称";
	m_WZCheck << "设计数量" << "上报数量" << "单价";
	//m_SGCheck << "序号" << "施工上报数量" << "施工上报单价";
	//m_JSCheck << "序号" << "结算数量" << "结算总价";
	//m_WZBCheck << "序号" << "价格" << "不含税价";
}

ExcelProcess::~ExcelProcess()
{

}

#include <QDate>
int g_DayNum = 35;
bool ExcelProcess::FileData(const QString strFileName, QStringList& arrField, QList<QStringList>& arrFileData, QString strFlag)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return false;
	}

	QList<int> arrCnt;
	bool bSuccess = false;
	Book* book;
	if (strFileName.right(3) == "xls")
	{
		book = xlCreateBookA();
	}
	else
	{
		book = xlCreateXMLBook();
	}

	QList<QStringList> arrTmp;
	const char * x = "Halil Kural";
	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	if(book->load(strFileName.toLocal8Bit().data()))
	{
		Sheet* sheet = book->getSheet(0);
		if(sheet)
		{
			bSuccess = true;

			int iFirstCol = sheet->firstCol();
			int iLastCol = sheet->lastCol();
			int iFirstRow = sheet->firstRow();
			int iLastRow = sheet->lastRow(); 
			for(int i = iFirstCol; i < iLastCol; i++)
			{
				const char* cColName = sheet->readStr(0, i);
				if(cColName)
				{
					//控制读取哪些字段
					QString strColName = QString::fromLocal8Bit(cColName);
					if(strFlag == QString::fromLocal8Bit("@WZ")|| strFlag == QString::fromLocal8Bit("@SG")|| strFlag == QString::fromLocal8Bit("@JS"))
					{
						if(!m_WZFlags.contains(strColName))
						{
							continue;
						}
					}
					else if (strFlag == QString::fromLocal8Bit("@WZB"))
					{
						if (!m_WZBFlags.contains(strColName))
						{
							continue;
						}
					}

					QStringList arrColInfo;
					for(int j = iFirstRow+1; j < iLastRow; j++)
					{
						if(sheet->cellType(j,i) == CELLTYPE_EMPTY)
						{
							if(m_WZCheck.contains(strColName))
							{
								//arrColInfo << QString::fromLocal8Bit("-1@@@@@");
								arrColInfo << QString::fromLocal8Bit("");
							}
							else if (m_XMCheck.contains(strColName))
							{
								arrColInfo << QString::fromLocal8Bit("-1@@@@@");
							}
							else
							{
								arrColInfo << QString::fromLocal8Bit("");
							}
						}
						else if(sheet->cellType(j,i) == CELLTYPE_NUMBER)
						{
							double fTemp = sheet->readNum(j,i);
							if(fTemp == 0 && (m_WZCheck.contains(strColName)))
							{
								if(m_WZCheck.contains(strColName))
								{
									//arrColInfo << QString::fromLocal8Bit("-1@@@@@");
									arrColInfo << QString::fromLocal8Bit("");
								}
								else
								{
									arrColInfo << QString::fromLocal8Bit("");
								}
								
							}
							else
							{
								QVariant v(fTemp);
								arrColInfo << v.toString();
							}
						}
						else
						{							
							const char* cInfo = sheet->readStr(j,i);
							if(cInfo == NULL) 
							{
								//if(m_WZCheck.contains(strColName)) arrColInfo << "-1@@@@@";
								if(m_WZCheck.contains(strColName)) arrColInfo << "";
								else arrColInfo << "";
							}
							else if(m_WZCheck.contains(strColName))
							{
								//arrColInfo << QString::fromLocal8Bit("-1@@@@@");
								arrColInfo << QString::fromLocal8Bit(cInfo);
							}
							else
							{
								arrColInfo << QString::fromLocal8Bit(cInfo);
							}
						}
					}

					arrField << strColName;
					arrFileData << arrColInfo;
					arrCnt << arrColInfo.size();

					
				}
			}
		}
		else
		{
			QMessageBox::information(NULL, QString::fromLocal8Bit("提示"),QString::fromLocal8Bit("Excel加载Sheet失败！") ,QMessageBox::Ok);
		}
	}
	else
	{
		QMessageBox::information(NULL, QString::fromLocal8Bit("提示"),QString::fromLocal8Bit("文件加载失败！") ,QMessageBox::Ok);
	}

	book->release();

	if (arrFileData.size() > 0)
	{
		for (int i = 0; i < arrField.size(); i++)
		{
			QStringList& arr = arrFileData[i];
			for (int j = arrFileData[0].size()-1; j >= 0; j--)
			{
				if (arr[j] == "-1@@@@@")
				{//删除此行数据
					for (int k = 0; k < arrField.size(); k++)
					{
						arrFileData[k].removeAt(j);
					}
				}
			}
		}
	}
	

	return bSuccess;
}
#include <QString>

bool ExcelProcess::SaveToFile(const QMap<QString, QStringList>& arrData, QString strFileName)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return false;
	}

	bool bSuccess = false;
	Book* book = xlCreateXMLBook();
	const char * x = "Halil Kural";
	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	Sheet* sheet;
	if(book->load(strFileName.toLocal8Bit().data()))
	{
		sheet = book->getSheet(0);
		if(sheet)
		{
			bSuccess = true;

			int iIndex = 0;
			QMap<QString, QStringList>::Iterator item = arrData.begin();
			for (;item!=arrData.end(); item++)
			{
				iIndex++;
				for (int i = 0; i < item.value().size()-1; i++)
				{
					sheet->writeStr(iIndex, i, item.value().at(i).toLocal8Bit().data());
				}
			}
		}

		book->save(strFileName.toLocal8Bit().data());
	}
	book->release();

	return bSuccess;
}

bool ExcelProcess::SaveToFile(QList<QStringList> arrData, const QString strFileName, QList<int> arrNumIndex, int iSheetIndex)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return false;
	}

	bool bSuccess = false;
	Book* book = xlCreateXMLBook();
	const char * x = "Halil Kural";
	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	Sheet* sheet;
	if(book->load(strFileName.toLocal8Bit().data()))
	{
		sheet = book->getSheet(iSheetIndex);
		if(sheet)
		{
			bSuccess = true;

			int iIndex = 0;
			QList<QStringList>::Iterator item = arrData.begin();
			for (;item!=arrData.end(); item++)
			{
				iIndex++;
				Format* format = book->addFormat();

				if (iSheetIndex == 1 && item->at(11).isEmpty())
				{//如果物资清册未查询到数据，则标红显示
					format->setFillPattern(FILLPATTERN_SOLID);
					format->setPatternForegroundColor(COLOR_RED);
				}

				if (iSheetIndex == 0 && item->at(7).isEmpty())
				{//如果物资清册未查询到数据，则标红显示
					format->setFillPattern(FILLPATTERN_SOLID);
					format->setPatternForegroundColor(COLOR_RED);
				}

				for (int i = 0; i < item->size(); i++)
				{
					if (arrNumIndex.contains(i) && !item->at(i).isEmpty())
					{
						/*QString str = item->at(i);
						sheet->writeNum(iIndex, i, item->at(i).toDouble(),format);*/
						QString str = QString::number(item->at(i).toDouble(),'f',4).replace(".0000","");
						sheet->writeStr(iIndex, i, str.toLocal8Bit().data(),format);
					}
					else
					{
						sheet->writeStr(iIndex, i, item->at(i).toLocal8Bit().data(), format);
					}
				}
			}
		}

		book->save(strFileName.toLocal8Bit().data());
	}
	book->release();

	return bSuccess;
}

//
bool ExcelProcess::SaveToFile(QList<ProjInfo> arrProInfo, const QString strFileName)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return false;
	}

	bool bSuccess = false;
	Book* book = xlCreateXMLBook();
	const char * x = "Halil Kural";
	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	Sheet* sheet;
	if(book->load(strFileName.toLocal8Bit().data()))
	{
		sheet = book->getSheet(0);
		if(sheet)
		{
			bSuccess = true;

			for (int i = 0; i < arrProInfo.size(); i++)
			{
				sheet->writeStr(i+1, 0, arrProInfo.at(i).ID);
				sheet->writeStr(i+1, 1, arrProInfo.at(i).Name);
				//sheet->writeStr(i+1, 2, QString::number(arrProInfo.at(i).iYear).toLocal8Bit().data());
				sheet->writeNum(i+1, 2, arrProInfo.at(i).iYear);
				sheet->writeStr(i+1, 3,arrProInfo.at(i).PC);
				sheet->writeStr(i+1, 4, arrProInfo.at(i).type);
				sheet->writeStr(i+1, 5, arrProInfo.at(i).CoLtd);
				sheet->writeStr(i+1, 6, QString::number(arrProInfo.at(i).fTZ,'f',4).toLocal8Bit().data());
				//sheet->writeNum(i+1, 6, arrProInfo.at(i).fTZ);
			}
		}

		book->save(strFileName.toLocal8Bit().data());
	}
	book->release();

	return bSuccess;
}

bool ExcelProcess::SaveToFile(ProjInfo info, QStringList arrField, QList<QStringList> arrFileData, const QString strFileName, const QList<int> arrNumIndex, QList<double> arrZJ)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return false;
	}

	bool bSuccess = false;
	Book* book = xlCreateXMLBook();
	const char * x = "Halil Kural";
	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	Sheet* sheet;
	if(book->load(strFileName.toLocal8Bit().data()))
	{
		sheet = book->getSheet(0);
		if(sheet)
		{
			bSuccess = true;

			//写入项目相关信息
			sheet->writeStr(0, 1,	info.ID);
			sheet->writeStr(0, 6,	info.Name);
			sheet->writeStr(0, 10,	info.CoLtd);
			sheet->writeNum(1, 1,	info.iYear);
			sheet->writeStr(1, 4,	info.PC);
			sheet->writeStr(1, 6,	info.type);
			sheet->writeNum(1, 10,	info.fTZ);

			for (int i = 0; i < arrField.size(); i++)
			{
				for (int j = 5; j < arrFileData[i].size()+5; j++)
				{
					if (arrNumIndex.contains(i) && !arrFileData[i].at(j-5).isEmpty())
					//if (arrNumIndex.contains(i))
					{//指定列按数字存储
						double f = arrFileData[i].at(j-5).toDouble();

						f = QString::number(f,'f',4).replace(".0000","").toDouble();
						sheet->writeNum(j, i,  f);
						//sheet->writeStr(j, i,  (const char*)(QString::number(f,'f',4).replace(".0000","").toLocal8Bit().data()));
					}
					else
					{
						sheet->writeStr(j, i, (const char*)(arrFileData[i].at(j-5).toLocal8Bit().data()));
					}					
				}
			}
		}

		//写入总价数据
		if (arrZJ.size() == 5)
		{
			sheet->writeStr(3, 7,	 QString("总价: %1").	arg(arrZJ[0],0,'f',4).toLocal8Bit().data());
			sheet->writeStr(3, 9,	 QString("总价: %1").	arg(arrZJ[1],0,'f',4).toLocal8Bit().data());
			sheet->writeStr(3, 12,  QString("总价: %1").	arg(arrZJ[2],0,'f',4).toLocal8Bit().data());
			sheet->writeStr(3, 13,  QString("总价: %1").	arg(arrZJ[3],0,'f',4).toLocal8Bit().data());
			sheet->writeStr(3, 15,  QString("总价: %1").	arg(arrZJ[4],0,'f',4).toLocal8Bit().data());
		}

		book->save(strFileName.toLocal8Bit().data());
	}
	book->release();

	return bSuccess;
}

//导出查询项目汇总信息
bool ExcelProcess::SaveToFile(QStringList arrField, QList<QStringList> arrFileData, Book* book, Sheet* sheet, const QList<int> arrNumIndex, QList<double> arrZJ, int iCurRow)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return false;
	}

	if(sheet)
	{
		//写入表头
		//Format* fmt = book->addFormat();
		//fmt->setFillPattern(FILLPATTERN_SOLID);
		//fmt->setPatternForegroundColor(COLOR_GREEN);
		//fmt->setBorder(BORDERSTYLE_THIN);
		//fmt->setAlignH(ALIGNH_CENTER);
		//fmt->setAlignV(ALIGNV_CENTER);
		//for (int i = 0; i < arrField.size(); i++)
		//{
		//	sheet->writeStr(iCurRow, i, arrField[i].toLocal8Bit().data(), fmt);
		//	sheet->writeStr(iCurRow+1, i, arrField[i].toLocal8Bit().data(), fmt);
		//	sheet->writeStr(iCurRow+2, i, arrField[i].toLocal8Bit().data(), fmt);
		//}

		////合并相关单元格
		//for (int i = 0; i < 9; i++)
		//{
		//	sheet->setMerge(iCurRow, iCurRow+2,i,i);
		//}

		////物资表
		//sheet->writeStr(iCurRow, 9, "协议物资表");
		//sheet->writeStr(iCurRow+1, 9, QString::number(arrZJ[0],'f',3).toLocal8Bit().data());
		//sheet->setMerge(iCurRow,  iCurRow,9, 10);
		//sheet->setMerge(iCurRow+1, iCurRow+1,9,  10);

		////物资清册
		//sheet->writeStr(iCurRow, 11, "物资清册");
		//sheet->writeStr(iCurRow+1, 11, QString::number(arrZJ[1],'f',3).toLocal8Bit().data());
		//sheet->setMerge(iCurRow,  iCurRow,11, 13);
		//sheet->setMerge(iCurRow+1,  iCurRow+1,11, 13);

		////施工清册
		//sheet->writeStr(iCurRow, 14, "施工清册");
		//sheet->writeStr(iCurRow+1, 14, QString::number(arrZJ[2],'f',3).toLocal8Bit().data());

		////结算清册
		//sheet->writeStr(iCurRow, 15, "结算清册");
		//sheet->writeStr(iCurRow+1, 15, QString::number(arrZJ[3],'f',3).toLocal8Bit().data());
		//sheet->setMerge(iCurRow,  iCurRow,15, 16);
		//sheet->setMerge(iCurRow+1, iCurRow+1,15,  16);

		////ERP清册
		//sheet->writeStr(iCurRow, 17, "ERP清册");
		//sheet->writeStr(iCurRow+1, 17, QString::number(arrZJ[4],'f',3).toLocal8Bit().data());
		//sheet->setMerge(iCurRow,  iCurRow,17, 19);
		//sheet->setMerge(iCurRow+1,  iCurRow+1,17, 19);

		//iCurRow += 3;


		for (int i = 0; i < arrField.size(); i++)
		{
			for (int j = 0; j < arrFileData[i].size(); j++)
			{
				if (arrNumIndex.contains(i) && !arrFileData[i].at(j).isEmpty())
					//if (arrNumIndex.contains(i))
				{//指定列按数字存储
					//sheet->writeNum(j+iCurRow, i, (arrFileData[i].at(j).toDouble()));
					double f = arrFileData[i].at(j).toDouble();

					f = QString::number(f,'f',4).replace(".0000","").toDouble();
					sheet->writeNum(j+iCurRow, i,  f);
					//sheet->writeStr(j+iCurRow, i, (const char*)());
				}
				else
				{
					sheet->writeStr(j+iCurRow, i, (const char*)(arrFileData[i].at(j).toLocal8Bit().data()));
				}					
			}
		}

		//数据比较标记（单个项目的物资清册单价跟物资表的单价进行对比，单价超过或低于标准价的20%进行标红提示；
		//施工清册的上报数量和物资清册上报数量进行对比，有数量不一致的进行提示 ）


		for (int j = 0; j < arrFileData[0].size(); j++)
		{
			QList<int> arrII;
			for (int i = 0; i < arrField.size(); i++) arrII<<arrFileData[i].size();
			{
				//单价
				double f1 = arrFileData[10].at(j).toDouble();
				double f2 = arrFileData[13].at(j).toDouble();
				if (f1== 0) f1=1;

				f1 = fabs(f1-f2)/f1;
				if (f1>0.2)
				{//标记红色
					Format* format = book->addFormat();
					format->setFillPattern(FILLPATTERN_SOLID);
					format->setPatternForegroundColor(COLOR_RED);
					sheet->writeNum(j+iCurRow, 10, (arrFileData[10].at(j).toDouble()), format);
					sheet->writeNum(j+iCurRow, 13, (arrFileData[13].at(j).toDouble()), format);
				}

				//数量
				if (arrFileData[12].at(j).toDouble() != arrFileData[14].at(j).toDouble())
				{//标记黄色
					Format* format = book->addFormat();
					format->setFillPattern(FILLPATTERN_SOLID);
					format->setPatternForegroundColor(COLOR_YELLOW);
					sheet->writeNum(j+iCurRow,12, (arrFileData[12].at(j).toDouble()), format);
					sheet->writeNum(j+iCurRow, 14, (arrFileData[14].at(j).toDouble()), format);
				}
			}
		}
	}

	return true;
}

//bool ExcelProcess::SaveToFile(QStringList arrField, QList<QStringList> arrFileData, const QString strFileName, const QList<int> arrNumIndex, QList<double> arrZJ, int iCurRow)
//{
//
//	bool bSuccess = false;
//	Book* book = xlCreateXMLBook();
//	const char * x = "Halil Kural";
//	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
//	book->setKey(x, y);
//	Sheet* sheet;
//	if(book->load(strFileName.toLocal8Bit().data()))
//	{
//		sheet = book->getSheet(0);
//		if(sheet)
//		{
//			bSuccess = true;
//
//			//写入表头
//			Format* fmt = book->addFormat();
//			fmt->setFillPattern(FILLPATTERN_SOLID);
//			fmt->setPatternForegroundColor(COLOR_GREEN);
//			fmt->setBorder(BORDERSTYLE_THIN);
//			fmt->setAlignH(ALIGNH_CENTER);
//			fmt->setAlignV(ALIGNV_CENTER);
//			for (int i = 0; i < arrField.size(); i++)
//			{
//				sheet->writeStr(iCurRow, i, arrField[i].toLocal8Bit().data(), fmt);
//				sheet->writeStr(iCurRow+1, i, arrField[i].toLocal8Bit().data(), fmt);
//				sheet->writeStr(iCurRow+2, i, arrField[i].toLocal8Bit().data(), fmt);
//			}
//
//			//合并相关单元格
//			for (int i = 0; i < 9; i++)
//			{
//				sheet->setMerge(iCurRow, iCurRow+2,i,i);
//			}
//
//			//物资表
//			sheet->writeStr(iCurRow, 9, "协议物资表");
//			sheet->writeStr(iCurRow+1, 9, QString::number(arrZJ[0],'f',3).toLocal8Bit().data());
//			sheet->setMerge(iCurRow,  iCurRow,9, 10);
//			sheet->setMerge(iCurRow+1, iCurRow+1,9,  10);
//
//			//物资清册
//			sheet->writeStr(iCurRow, 11, "物资清册");
//			sheet->writeStr(iCurRow+1, 11, QString::number(arrZJ[1],'f',3).toLocal8Bit().data());
//			sheet->setMerge(iCurRow,  iCurRow,11, 13);
//			sheet->setMerge(iCurRow+1,  iCurRow+1,11, 13);
//
//			//施工清册
//			sheet->writeStr(iCurRow, 14, "施工清册");
//			sheet->writeStr(iCurRow+1, 14, QString::number(arrZJ[2],'f',3).toLocal8Bit().data());
//
//			//结算清册
//			sheet->writeStr(iCurRow, 15, "结算清册");
//			sheet->writeStr(iCurRow+1, 15, QString::number(arrZJ[3],'f',3).toLocal8Bit().data());
//			sheet->setMerge(iCurRow,  iCurRow,15, 16);
//			sheet->setMerge(iCurRow+1, iCurRow+1,15,  16);
//
//			//ERP清册
//			sheet->writeStr(iCurRow, 17, "ERP清册");
//			sheet->writeStr(iCurRow+1, 17, QString::number(arrZJ[4],'f',3).toLocal8Bit().data());
//			sheet->setMerge(iCurRow,  iCurRow,17, 19);
//			sheet->setMerge(iCurRow+1,  iCurRow+1,17, 19);
//
//			iCurRow += 3;
//
//
//			for (int i = 0; i < arrField.size(); i++)
//			{
//				for (int j = 0; j < arrFileData[i].size(); j++)
//				{
//					if (arrNumIndex.contains(i) && !arrFileData[i].at(j).isEmpty())
//						//if (arrNumIndex.contains(i))
//					{//指定列按数字存储
//						sheet->writeNum(j+iCurRow, i, (arrFileData[i].at(j).toDouble()));
//					}
//					else
//					{
//						sheet->writeStr(j+iCurRow, i, (const char*)(arrFileData[i].at(j).toLocal8Bit().data()));
//					}					
//				}
//			}
//
//			//数据比较标记（单个项目的物资清册单价跟物资表的单价进行对比，单价超过或低于标准价的20%进行标红提示；
//			//施工清册的上报数量和物资清册上报数量进行对比，有数量不一致的进行提示 ）
//
//
//			for (int j = 0; j < arrFileData[0].size(); j++)
//			{
//				QList<int> arrII;
//				for (int i = 0; i < arrField.size(); i++) arrII<<arrFileData[i].size();
//				{
//					//单价
//					double f1 = arrFileData[10].at(j).toDouble();
//					double f2 = arrFileData[13].at(j).toDouble();
//					if (f1== 0) f1=1;
//
//					f1 = fabs(f1-f2)/f1;
//					if (f1>0.2)
//					{//标记红色
//						Format* format = book->addFormat();
//						format->setFillPattern(FILLPATTERN_SOLID);
//						format->setPatternForegroundColor(COLOR_RED);
//						sheet->writeNum(j+iCurRow, 13, (arrFileData[13].at(j).toDouble()), format);
//					}
//
//					//数量
//					if (arrFileData[12].at(j).toDouble() != arrFileData[14].at(j).toDouble())
//					{//标记黄色
//						Format* format = book->addFormat();
//						format->setFillPattern(FILLPATTERN_SOLID);
//						format->setPatternForegroundColor(COLOR_YELLOW);
//						sheet->writeNum(j+iCurRow,12, (arrFileData[12].at(j).toDouble()), format);
//						sheet->writeNum(j+iCurRow, 14, (arrFileData[14].at(j).toDouble()), format);
//					}
//				}
//			}
//		}
//
//		book->save(strFileName.toLocal8Bit().data());
//	}
//	book->release();
//
//	return bSuccess;
//}
//导出ERP本次处理结果数据
bool ExcelProcess::SaveToFile(QList<QStringList> arrData, QString strFileName)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return false;
	}

	bool bSuccess = false;
	Book* book = xlCreateXMLBook();
	const char * x = "Halil Kural";
	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	Sheet* sheet;
	if(book->load(strFileName.toLocal8Bit().data()))
	{
		sheet = book->getSheet(0);
		if(sheet)
		{
			bSuccess = true;

			int iIndex = 0;
			QList<QStringList>::Iterator item = arrData.begin();
			for (;item!=arrData.end(); item++)
			{
				iIndex++;
				for (int i = 0; i < item->size(); i++)
				{
					sheet->writeStr(iIndex, i, item->at(i).toLocal8Bit().data());
				}
			}
		}

		book->save(strFileName.toLocal8Bit().data());
	}
	book->release();

	return bSuccess;
}

//导出查询项目的铁附件汇总数据
bool ExcelProcess::SaveToFile(QString strFileName, QList<QStringList> arrData1, QList<QStringList> arrData2)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return false;
	}

	bool bSuccess = false;
	Book* book = xlCreateXMLBook();
	const char * x = "Halil Kural";
	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	Sheet* sheet;
	if(book->load(strFileName.toLocal8Bit().data()))
	{
		sheet = book->getSheet(0);
		if(sheet)
		{
			bSuccess = true;

			int iIndex = 0;
			QList<QStringList>::Iterator item = arrData1.begin();
			for (;item!=arrData1.end(); item++)
			{
				iIndex++;
				int iColumn = 0;//列
				sheet->writeStr(iIndex, 0, QString::number(iIndex).toLocal8Bit().data());
				for (int i = 0; i < item->size(); i++)
				{
					if(i == 6) iColumn++;
					if(i == 7) iColumn-=2;
					iColumn++;
					sheet->writeStr(iIndex, iColumn, item->at(i).toLocal8Bit().data());
				}
			}
		}

		sheet = book->getSheet(1);
		if(sheet)
		{
			bSuccess = true;

			int iIndex = 0;
			QList<QStringList>::Iterator item = arrData2.begin();
			for (;item!=arrData2.end(); item++)
			{
				iIndex++;
				int iColumn = 0;//列
				sheet->writeStr(iIndex, iColumn, QString::number(iIndex).toLocal8Bit().data());
				for (int i = 3; i < item->size(); i++)
				{
					if(i == 6) iColumn++;
					if(i == 7) iColumn-=2;
					iColumn++;	
					sheet->writeStr(iIndex, iColumn, item->at(i).toLocal8Bit().data());
				}
			}
		}

		book->save(strFileName.toLocal8Bit().data());
	}
	book->release();

	return bSuccess;
}

//将表格中的数据进行分析，将不同地方用颜色进行提示
void ExcelProcess::ExcelDataCompare(const QString strFileName)
{
	if (abs(QDate::currentDate().daysTo(QDate(2019,3,1))) > g_DayNum)
	{
		return ;
	}

	bool bSuccess = false;
	Book* book = xlCreateXMLBook();
	const char * x = "Halil Kural";
	const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	Sheet* sheet;
	if(book->load(strFileName.toLocal8Bit().data()))
	{
		sheet = book->getSheet(0);
		if(sheet)
		{
			bSuccess = true;

			//施工阶段数量和物资清册的有区别用红色背景提醒
			int iWZQCSlIndex = 11;
			int iSGQCSlIndex = 14;
			for(int i = 4; i < sheet->lastRow()-1; i++)
			{
				QString strWZQCSl = sheet->readStr(i,iWZQCSlIndex);
				QString strSGQCSl = sheet->readStr(i,iSGQCSlIndex);
				if((strWZQCSl != QString::fromLocal8Bit("")) && (strSGQCSl != QString::fromLocal8Bit("")) && (strWZQCSl != strSGQCSl))
				{
					Format* format = book->addFormat();
					format->setFillPattern(FILLPATTERN_SOLID);
					format->setPatternForegroundColor(COLOR_RED);
					sheet->writeStr(i,iWZQCSlIndex,(const char*)strWZQCSl.toLocal8Bit().data(),format);
					sheet->writeStr(i,iSGQCSlIndex,(const char*)strSGQCSl.toLocal8Bit().data(),format);
				}
			}

			//施工阶段价格与选择的协议物资价格有区别也用红色提醒
			int iWZQCJgIndex = 12;
			int iWZQCZjIndex = 13;
			int iSGQCJgIndex = 15;
			for(int i = 4; i < sheet->lastRow()-1; i++)
			{
				QString strWZQCJg = sheet->readStr(i,iWZQCJgIndex);
				QString strSGQCJg = sheet->readStr(i,iSGQCJgIndex);
				if((strWZQCJg != QString::fromLocal8Bit("")) && (strSGQCJg != QString::fromLocal8Bit("")) && (strWZQCJg != strSGQCJg))
				{
					Format* format = book->addFormat();
					format->setFillPattern(FILLPATTERN_SOLID);
					format->setPatternForegroundColor(COLOR_RED);
					sheet->writeStr(i,iWZQCJgIndex,(const char*)strWZQCJg.toLocal8Bit().data(),format);
					sheet->writeStr(i,iSGQCJgIndex,(const char*)strSGQCJg.toLocal8Bit().data(),format);
				}
			}

			//施工清册上如果比物资清册多出材料用黄色背景提示
			for(int i = 4; i < sheet->lastRow()-1; i++)
			{
				QString strWZQCSl = sheet->readStr(i,iWZQCSlIndex);
				QString strSGQCSl = sheet->readStr(i,iSGQCSlIndex);
				QString strWZQCJg = sheet->readStr(i,iWZQCJgIndex);
				QString strWZQCZj = sheet->readStr(i,iWZQCZjIndex);
				QString strSGQCJg = sheet->readStr(i,iSGQCJgIndex);
				if((strWZQCSl == QString::fromLocal8Bit("")) && (strWZQCJg == QString::fromLocal8Bit("")) && (strSGQCSl != QString::fromLocal8Bit("")) && (strSGQCJg != QString::fromLocal8Bit("")))
				{
					Format* format = book->addFormat();
					format->setFillPattern(FILLPATTERN_SOLID);
					format->setPatternForegroundColor(COLOR_YELLOW);
					sheet->writeStr(i,iWZQCSlIndex,(const char*)strWZQCSl.toLocal8Bit().data(),format);
					sheet->writeStr(i,iWZQCJgIndex,(const char*)strWZQCJg.toLocal8Bit().data(),format);
					sheet->writeStr(i,iWZQCZjIndex,(const char*)strWZQCZj.toLocal8Bit().data(),format);
					sheet->writeStr(i,iSGQCSlIndex,(const char*)strSGQCSl.toLocal8Bit().data(),format);
					sheet->writeStr(i,iSGQCJgIndex,(const char*)strSGQCJg.toLocal8Bit().data(),format);
					for(int k = 0; k < 18; k++)
					{
						const char* strTemp = sheet->readStr(i,k);
						sheet->writeStr(i,k,strTemp,format);
					}
				}
			}

			//结算清册上如果和施工清册不同的材料用红色背景提示
			int iJSQCSlIndex = 16;
			int iJSQCZjIndex = 17;
			for(int i = 4; i < sheet->lastRow()-1; i++)
			{
				QString strSGQCSl = sheet->readStr(i,iSGQCSlIndex);
				QString strJSQCSl = sheet->readStr(i,iJSQCSlIndex);
				QString strSGQCJg = sheet->readStr(i,iSGQCJgIndex);
				QString strJSQCZj = sheet->readStr(i,iJSQCZjIndex);
				if((strSGQCSl == QString::fromLocal8Bit("")) && (strSGQCJg == QString::fromLocal8Bit("")) && (strJSQCSl != QString::fromLocal8Bit("")) && (strJSQCZj != QString::fromLocal8Bit("")))
				{
					Format* format = book->addFormat();
					format->setFillPattern(FILLPATTERN_SOLID);
					format->setPatternForegroundColor(COLOR_RED);
					sheet->writeStr(i,iSGQCSlIndex,(const char*)strSGQCSl.toLocal8Bit().data(),format);
					sheet->writeStr(i,iSGQCJgIndex,(const char*)strSGQCJg.toLocal8Bit().data(),format);
					sheet->writeStr(i,iJSQCSlIndex,(const char*)strJSQCSl.toLocal8Bit().data(),format);
					sheet->writeStr(i,iJSQCZjIndex,(const char*)strJSQCZj.toLocal8Bit().data(),format);
					for(int k = 0; k < 18; k++)
					{
						const char* strTemp = sheet->readStr(i,k);
						sheet->writeStr(i,k,strTemp,format);
					}
				}
			}
		}

		book->save(strFileName.toLocal8Bit().data());
	}

	book->release();
}


