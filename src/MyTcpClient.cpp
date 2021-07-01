//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives.
//
// Source code and executable can be downloaded from
//  https://sourceforge.net/projects/mbox-viewer/  and
//  https://github.com/eneam/mboxviewer
//
//  Copyright(C) 2019  Enea Mansutti, Zbigniew Minciel
//
//  This program is free software; you can redistribute it and/or modify
//  it under the terms of the version 3 of GNU Affero General Public License
//  as published by the Free Software Foundation; 
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//


#include "stdafx.h"
#include "MyTcpClient.h"

CString GetLastErrorAsString();

MyTcpClient::MyTcpClient(int port)
{
	m_TcpPort = port;
}

MyTcpClient::~MyTcpClient()
{
}

int  MyTcpClient::Connect()
{
	m_connected = s.Connect("127.0.0.1", m_TcpPort);
	return 0;
}

int  MyTcpClient::Close()
{
	s.Close();
	return 0;
}

int  MyTcpClient::Send(CString &msg)
{
	int length = s.Send(msg, msg.GetLength());
	return length;
}

int  MyTcpClient::ConnectSendClose(CString &msg, CString &errorText)
{
	char buff[2048];
	CSocket s;

	BOOL ret = s.Create();
	int error = 0;
	if (s.Connect("127.0.0.1", m_TcpPort))
	{
		int length = s.Send(msg, msg.GetLength());

		int i = s.Receive(buff, 2048);
		if (i >= 0)
			buff[i] = 0;
		else
			buff[0] = 0;
	}
	else
	{
		error = GetLastError();
		TRACE("Error: %d\n", error);
		CString msg = GetLastErrorAsString();
		TRACE("Error Msg: %s\n", msg);
	}
	s.Close();

	return error;
}