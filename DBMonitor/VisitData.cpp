// VisitData.cpp: implementation of the CVisitData class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "VisitData.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CVisitData::CVisitData()
{

}

CVisitData::~CVisitData()
{
	Uninitialize();
}

//��ʼ�����ݿ�ģ��
BOOL CVisitData::Initialize(LPCTSTR lpszDBServer, LPCTSTR lpszDBUserName,
							LPCTSTR lpszDBPassword, LPCTSTR lpszDBName, int nDBType) //nDBType=DB_TYPE_SQLSERVER
{
	try
	{
		Uninitialize();
		m_nDBType = nDBType;
		if (CString(lpszDBName).Right(4) == ".mdb")
			m_nDBType = DB_TYPE_ACCESS;

		m_strDBServer = lpszDBServer;		//���ݷ���������
		m_strDBUserName = lpszDBUserName;	//���ݿ������û�����
		m_strDBPassword = lpszDBPassword;	//���ݿ���������
		m_strDBName = lpszDBName;			//���ݿ�����

		m_pDBConn.CreateInstance(__uuidof(Connection));
		if(m_pDBConn == NULL)
		{
			return FALSE;
		}
		m_pDBRecordSet.CreateInstance(__uuidof(Recordset)); 
		if(m_pDBRecordSet == NULL)
		{
			return FALSE;
		}
		m_pStm.CreateInstance(__uuidof(Stream)); 
		if(m_pStm == NULL)
		{
			return FALSE;
		}

		//�������ݿ�
		CString strConn;
	//	access:strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=(DB file path like xxx.mdb);Jet OLEDB:Database Password=xxx";
	//	odbc:strConn = "Data Source=FdmNleDSN;User ID=sa";
		if (DB_TYPE_ACCESS == m_nDBType)
		{
			strConn = CString("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\\") + lpszDBName;//;Jet OLEDB:Database Password=2602";
//			CString strDBFile = ".\\";
//			strDBFile += lpszDBName;
//			if (! DoesTheFileExist(strDBFile))
//				CopyFile(".\\backnull.mdb", strDBFile, TRUE);
		}
		else if (DB_TYPE_SQLSERVER == m_nDBType)
		{
			if(m_strDBServer.Compare(DBSVRIP)==0)
				strConn = "Provider=SQLOLEDB;Data Source=" + m_strDBServer + ",1433;User ID=" + m_strDBUserName + ";Password=" + m_strDBPassword;
			else
				strConn = "Provider=SQLOLEDB;Server=" + m_strDBServer + ";User ID=" + m_strDBUserName + ";Password=" + m_strDBPassword;
		}
		BSTR bstrConn = strConn.AllocSysString();
		m_pDBConn->Open(bstrConn, "", "", adConnectUnspecified);
		::SysFreeString(bstrConn);
		if (adStateOpen != m_pDBConn->GetState())
			return FALSE;

		if (DB_TYPE_SQLSERVER == m_nDBType)
		{
			BSTR bstrDBName = m_strDBName.AllocSysString();
			m_pDBConn->PutDefaultDatabase(bstrDBName);
			::SysFreeString(bstrDBName);
		}

		return TRUE;
	}
	catch(_com_error &e)
	{
		DWORD dwErr = GetLastError();
		ASSERT(FALSE);
		AfxMessageBox(e.Description());
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//����ʼ�����ݿ�ģ��
BOOL CVisitData::Uninitialize()
{
	try
	{
		//�رռ�¼��,�ͷ���Դ
		if(m_pDBRecordSet != NULL)
		{
			if (adStateOpen == m_pDBRecordSet->GetState())
				m_pDBRecordSet->Close ();
			m_pDBRecordSet.Release();
		}
		// �Ͽ����ݿ�����,�ͷ���Դ
		if(m_pDBConn != NULL)
		{
			if (adStateOpen == m_pDBConn->GetState())
				m_pDBConn->Close ();
			m_pDBConn.Release();
		}
		// �ر�������,�ͷ���Դ
		if(m_pStm != NULL)
		{
			if (adStateOpen == m_pStm->GetState())
				m_pStm->Close ();
			m_pStm.Release();
		}

		return TRUE;
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//��ȡ��ѯ����ļ�¼��Ŀ
long CVisitData::GetRecordCount()
{
//	try
//	{
//		if (adStateOpen == m_pDBRecordSet->GetState ())
//		{
//			return m_pDBRecordSet->GetRecordCount();
//		}
//	}
//	catch(_com_error &e)
//	{
//		ASSERT(FALSE);
//		TRACE("%s\n",e.Description());
//	}
//	catch(...)
//	{
//		ASSERT(FALSE);
//		TRACE("exception !\n");
//	}
// 	return 0;
	return m_nRecordCount;
}

//����ǰ��¼ָ������ƶ�һ����¼
BOOL CVisitData::RecordSetMoveNext()
{
	try
	{
		if (m_pDBRecordSet != NULL)
			if (adStateOpen == m_pDBRecordSet->GetState ())
			{
				m_pDBRecordSet->MoveNext();
				return TRUE;
			}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//����ǰ��¼ָ����ǰ�ƶ�һ����¼
BOOL CVisitData::RecordSetMovePrev()
{
	try
	{
		if (m_pDBRecordSet != NULL)
			if (adStateOpen == m_pDBRecordSet->GetState ())
			{
				m_pDBRecordSet->MovePrevious();
				return TRUE;
			}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//�ƶ���ǰ��¼ָ��
BOOL CVisitData::RecordSetMove(long lStepSize)
{
	try
	{
		if (m_pDBRecordSet != NULL)
			if (adStateOpen == m_pDBRecordSet->GetState ())
			{
				m_pDBRecordSet->Move(lStepSize);
				return TRUE;
			}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//�����һ����¼�趨Ϊ��¼���ĵ�ǰ��¼
BOOL CVisitData::RecordSetMoveLast()
{
	try
	{
		if (m_pDBRecordSet != NULL)
			if (adStateOpen == m_pDBRecordSet->GetState ())
			{
				m_pDBRecordSet->MoveLast();
				return TRUE;
			}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//����һ����¼�趨Ϊ��¼���ĵ�ǰ��¼
BOOL CVisitData::RecordSetMoveFirst()
{
	try
	{
		if (m_pDBRecordSet != NULL)
			if (adStateOpen == m_pDBRecordSet->GetState ())
			{
				m_pDBRecordSet->MoveFirst();
				return TRUE;
			}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//��ȡ��ǰ��¼�Ƿ��Ǽ�¼�������һ����¼
BOOL CVisitData::IsRecordSetEOF()
{
	try
	{
		if (m_pDBRecordSet != NULL)
			if (adStateOpen == m_pDBRecordSet->GetState ())
				return m_pDBRecordSet->EndOfFile;
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return TRUE;
}

//��ȡ��ǰ��¼�Ƿ��Ǽ�¼���ĵ�һ����¼
BOOL CVisitData::IsRecordSetBOF()
{
	try
	{
		if (m_pDBRecordSet != NULL)
			if (adStateOpen == m_pDBRecordSet->GetState ())
				return m_pDBRecordSet->BOF;
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return TRUE;
}

//��ȡ��¼���Ƿ�Ϊ��
BOOL CVisitData::IsRecordSetEmpty()
{
	try
	{
		if (m_pDBRecordSet != NULL)
			if (adStateOpen == m_pDBRecordSet->GetState ())
				return (m_pDBRecordSet->BOF && m_pDBRecordSet->EndOfFile);
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return TRUE;
}

//��ɶ����ݵĲ�ѯ����
BOOL CVisitData::SearchData(LPCTSTR lpszSqlSel)
{
	try
	{
		CString strSqlSel(lpszSqlSel);
		if(strSqlSel == "" || m_pDBRecordSet == NULL)
			return FALSE;

		//���ԭ�еļ�¼��
		if (adStateOpen == m_pDBRecordSet->GetState())
		{
			m_pDBRecordSet->Close();
		}
		//��ȡ��¼��
//		CString strSqlForCount = strSqlSel;
//		strSqlForCount.MakeLower();
//		int nIndex = strSqlForCount.Find(" from ");
//		int nIndex2 = strSqlForCount.Find(" order ");
//		if (nIndex < 0)
//		{
//			ASSERT(FALSE);
//		}
//		else
//		{
//			if (nIndex2 < 0)
//				nIndex2 = strSqlSel.GetLength();
//			strSqlForCount = "select count(*) as TotalCount " + strSqlSel.Mid(nIndex, nIndex2 -  nIndex);
//			BSTR bstrSqlForCount = strSqlForCount.AllocSysString();
//			m_pDBRecordSet->Open(bstrSqlForCount, (IDispatch *)m_pDBConn, adOpenStatic, adLockOptimistic, adCmdUnknown);//ִ�в�ѯ
//			::SysFreeString(bstrSqlForCount);
//			if (adStateOpen == m_pDBRecordSet->GetState())
//			{
//				m_nRecordCount = GetFieldValue_Long("TotalCount");
//				m_pDBRecordSet->Close();
//			}
//		}
		//�����µļ�¼��adOpenKeyset adOpenDynamic adOpenStatic, adLockOptimistic, adCmdUnknown enmuCursorType
		BSTR bstrSqlSel = strSqlSel.AllocSysString();
		m_pDBRecordSet->Open(bstrSqlSel, (IDispatch *)m_pDBConn, adOpenDynamic, adLockOptimistic, adCmdUnknown);//ִ�в�ѯ
		::SysFreeString(bstrSqlSel);
		if (adStateOpen == m_pDBRecordSet->GetState())
			return TRUE;
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//��ȡ���ؽ����¼����ĳ���ֶε�ֵ
_variant_t CVisitData::GetFieldValueOfRecordSet(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue;
	try
	{
		CString strFieldName(lpszFieldName);
		if(strFieldName == "" || m_pDBRecordSet == NULL)
			return varFieldValue;

		if (adStateOpen == m_pDBRecordSet->GetState())
		{
			if ( ! (m_pDBRecordSet->BOF || m_pDBRecordSet->EndOfFile))//��¼����Ϊ��
			{
				BSTR bstrFieldName = strFieldName.AllocSysString();
				varFieldValue = m_pDBRecordSet->GetCollect(bstrFieldName);
				::SysFreeString(bstrFieldName);
			}
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return varFieldValue;
}

//��ȡ���ؽ����¼����ĳ���ֶε�ֵ��long��
long CVisitData::GetFieldValue_Long(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue = GetFieldValueOfRecordSet(lpszFieldName);
	if (varFieldValue.vt != VT_NULL && varFieldValue.vt != VT_EMPTY)
		return long(varFieldValue);
	else
		return 0;
}

//��ȡ���ؽ����¼����ĳ���ֶε�ֵ��short��
short CVisitData::GetFieldValue_Short(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue = GetFieldValueOfRecordSet(lpszFieldName);
	if (varFieldValue.vt != VT_NULL && varFieldValue.vt != VT_EMPTY)
		return short(varFieldValue);
	else
		return 0;
}

//��ȡ���ؽ����¼����ĳ���ֶε�ֵ��BYTE��
BYTE CVisitData::GetFieldValue_Byte(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue = GetFieldValueOfRecordSet(lpszFieldName);
	if (varFieldValue.vt != VT_NULL && varFieldValue.vt != VT_EMPTY)
		return BYTE(varFieldValue);
	else
		return 0;
}

//��ȡ���ؽ����¼����ĳ���ֶε�ֵ��BOOL��
BOOL CVisitData::GetFieldValue_Bool(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue = GetFieldValueOfRecordSet(lpszFieldName);
	if (varFieldValue.vt != VT_NULL && varFieldValue.vt != VT_EMPTY)
		return bool(varFieldValue);
	else
		return FALSE;
}

//��ȡ���ؽ����¼����ĳ���ֶε�ֵ��float��
float CVisitData::GetFieldValue_Float(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue = GetFieldValueOfRecordSet(lpszFieldName);
	if (varFieldValue.vt != VT_NULL && varFieldValue.vt != VT_EMPTY)
		return float(varFieldValue);
	else
		return 0;
}

//��ȡ���ؽ����¼����ĳ���ֶε�ֵ��double��
double CVisitData::GetFieldValue_Double(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue = GetFieldValueOfRecordSet(lpszFieldName);
	if (varFieldValue.vt != VT_NULL && varFieldValue.vt != VT_EMPTY)
		return double(varFieldValue);
	else
		return 0;
}

//��ȡ���ؽ����¼����ĳ���ֶε�ֵ��string��
CString CVisitData::GetFieldValue_String(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue = GetFieldValueOfRecordSet(lpszFieldName);
	if (varFieldValue.vt != VT_NULL && varFieldValue.vt != VT_EMPTY)
		return CString(varFieldValue.bstrVal);
	else
		return "";
}

//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��COleDateTime��
COleDateTime CVisitData::GetFieldValue_DateTime(LPCTSTR lpszFieldName)
{
	_variant_t varFieldValue = GetFieldValueOfRecordSet(lpszFieldName);
	if (varFieldValue.vt != VT_NULL && varFieldValue.vt != VT_EMPTY)
		return varFieldValue.date;
	else
		return COleDateTime::GetCurrentTime();
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ
BOOL CVisitData::SetFieldValueOfRecordSet(LPCTSTR lpszFieldName, _variant_t varValue)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			if (VARTYPE(VT_ARRAY | VT_UI1) == varValue.vt)
				m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = varValue;
			else
				m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->AppendChunk(varValue);
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ��long��
BOOL CVisitData::SetFieldValue_Long(LPCTSTR lpszFieldName, long lValue)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = lValue;
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ��short��
BOOL CVisitData::SetFieldValue_Short(LPCTSTR lpszFieldName, short nValue)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = nValue;
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ��BYTE��
BOOL CVisitData::SetFieldValue_Byte(LPCTSTR lpszFieldName, BYTE byValue)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = byValue;
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ��BOOL��
BOOL CVisitData::SetFieldValue_Bool(LPCTSTR lpszFieldName, BOOL bValue)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = (bool)bValue;
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ��float��
BOOL CVisitData::SetFieldValue_Float(LPCTSTR lpszFieldName, float fValue)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = fValue;
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ��double��
BOOL CVisitData::SetFieldValue_Double(LPCTSTR lpszFieldName, double dValue)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = dValue;
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ��string��
BOOL CVisitData::SetFieldValue_String(LPCTSTR lpszFieldName, LPCTSTR lpszValue)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = lpszValue;
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���õ�ǰ��¼����ĳ���ֶε�ֵ��COleDateTime��
BOOL CVisitData::SetFieldValue_DateTime(LPCTSTR lpszFieldName, COleDateTime otDateTime)
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->GetFields()->GetItem(lpszFieldName)->Value = otDateTime.m_dt;
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//���µ�ǰ��¼��
BOOL CVisitData::UpdateRecordSet()
{
	try
	{
		if (m_pDBRecordSet != NULL)
		{
			m_pDBRecordSet->Update();
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//ִ��SQL����
BOOL CVisitData::Execute(LPCTSTR lpszSql)
{
	try
	{
		CString strSql(lpszSql);
		if (strSql == "" || m_pDBConn == NULL)
			return FALSE;
		if (adStateOpen != m_pDBConn->GetState())
			return FALSE;

		//ִ�в���
		BSTR bstrSql = strSql.AllocSysString();
		m_pDBConn->Execute(bstrSql, NULL, adCmdUnknown);
		::SysFreeString(bstrSql);
		return TRUE;
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//�������ݣ���ȡ���Զ����ɵ�ID
int  CVisitData::InsertForID(LPCTSTR lpszSql)
{
	try
	{
		CString strSql(lpszSql);
		if (strSql == "" || m_pDBConn == NULL)
			return FALSE;
		if (adStateOpen != m_pDBConn->GetState())
			return FALSE;

		//ִ�в���
		if (DB_TYPE_SQLSERVER == m_nDBType)
		{
			strSql = "SET NOCOUNT ON;" + strSql + ";SELECT @@IDENTITY";
			BSTR bstrSql = strSql.AllocSysString();
			_RecordsetPtr pRecordSet = m_pDBConn->Execute(bstrSql, NULL, adCmdUnknown);
			_variant_t varID = pRecordSet->GetCollect(_variant_t((long)0));
			::SysFreeString(bstrSql);
			pRecordSet->Close();
			return (long)varID;
		}
		else if (DB_TYPE_ACCESS == m_nDBType)
		{
			BSTR bstrSql = strSql.AllocSysString();
			m_pDBConn->Execute(bstrSql, NULL, adCmdUnknown);
			::SysFreeString(bstrSql);
			strSql = "select max(ProblemID) as CurID from ProblemInfo";
			bstrSql = strSql.AllocSysString();
			_RecordsetPtr pRecordSet = m_pDBConn->Execute(bstrSql, NULL, adCmdUnknown);
			::SysFreeString(bstrSql);
			_variant_t varID = pRecordSet->GetCollect("CurID");
			pRecordSet->Close();
			return (long)varID;
		}
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return -1;
}

//��ʼ�������
BOOL CVisitData::BeginTrans()
{
	try
	{
		if (m_pDBConn == NULL)
			return FALSE;
		if (adStateOpen != m_pDBConn->GetState())
			return FALSE;
		m_pDBConn->BeginTrans();
		return TRUE;
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//�ύ����
BOOL CVisitData::CommitTrans()
{
	try
	{
		if (m_pDBConn == NULL)
			return FALSE;
		if (adStateOpen != m_pDBConn->GetState())
			return FALSE;
		HRESULT hr = m_pDBConn->CommitTrans();
		return TRUE;
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}

//����ִ��ʧ�ܣ��ع�
BOOL CVisitData::RollbackTrans()
{
	try
	{
		if (m_pDBConn == NULL)
			return FALSE;
		if (adStateOpen != m_pDBConn->GetState())
			return FALSE;
		m_pDBConn->RollbackTrans();
		return TRUE;
	}
	catch(_com_error &e)
	{
		ASSERT(FALSE);
		TRACE("%s\n",e.Description());
	}
	catch(...)
	{
		ASSERT(FALSE);
		TRACE("exception !\n");
	}
	return FALSE;
}
