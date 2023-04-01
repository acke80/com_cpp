#pragma once

#include <Windows.h>
#include <string>
#include <cassert>

namespace com
{

class ComPort
{
public:
    typedef HANDLE  port_t;
    typedef DCB     params_t;

private:
    port_t m_port{ nullptr };
    params_t m_params{ sizeof(params_t) };

    unsigned int m_com_index{ 0 };
    bool m_open{ false };

public:
    ComPort(unsigned int com_index) : m_com_index(com_index) 
    {
        setBaudRate(19200);
        setParity(0);
        setDataBits(8);
        setStopBits(1);
    }

    ~ComPort()
    {
        flush();
        
        if (m_open)
        {
            close();
        }
    }

    bool setBaudRate(unsigned int baud_rate)
    {
        if (!GetCommState(m_port, &m_params)) return false;

        m_params.BaudRate = baud_rate;
        
        return SetCommState(m_port, &m_params);
    }

    bool setParity(unsigned int parity)
    {
        if (!GetCommState(m_port, &m_params)) return false;

        m_params.Parity = parity;

        return SetCommState(m_port, &m_params);

    }

    bool setDataBits(unsigned int bits)
    {
        if (!GetCommState(m_port, &m_params)) return false;

        m_params.ByteSize = bits;

        return SetCommState(m_port, &m_params);
    }

    bool setStopBits(unsigned int bits)
    {
        if (!GetCommState(m_port, &m_params)) return false;

        m_params.StopBits = bits;

        return SetCommState(m_port, &m_params);
    }

    bool open()
    {
        port_t port;
        TCHAR com_name[100];
        wsprintf(com_name, TEXT("\\\\.\\COM%d"), m_com_index);
        port = CreateFile(com_name, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);

        if (port == INVALID_HANDLE_VALUE)
        {
            return false;
        }

        COMMTIMEOUTS timeouts = { 0 };
        timeouts.ReadIntervalTimeout            = 50;
        timeouts.ReadTotalTimeoutConstant       = 50;
        timeouts.ReadTotalTimeoutMultiplier     = 10;
        timeouts.WriteTotalTimeoutConstant      = 50;
        timeouts.WriteTotalTimeoutMultiplier    = 10;

        if (SetCommTimeouts(port, &timeouts) == FALSE || SetCommMask(port, EV_RXCHAR) == FALSE)
        {
            return false;
        }
        else
        {
            m_open = true;
            m_port = port;
        }
    }

    void close()
    {
        if (m_open)
        {
            CloseHandle(m_port);
            m_open = false;
        }
    }

    bool write(const std::string& data)
    {
        if (!m_open) return false;

        DWORD bytesWritten;
        bool status = WriteFile(
            m_port,
            data.c_str(),
            data.length(),
            &bytesWritten,
            NULL);

        return status;
    }

    bool read(char* data, unsigned int num_bytes)
    {
        if (!m_open) return false;

        DWORD dwEventMask;
        DWORD NoBytesRead;

        if (!WaitCommEvent(m_port, &dwEventMask, nullptr))
        {
            return false;
        }
        
        return ReadFile(m_port, data, num_bytes, &NoBytesRead, nullptr);
    }

    void flush()
    {
        if (!m_open) return;

        PurgeComm(m_port, PURGE_RXCLEAR);
        PurgeComm(m_port, PURGE_TXCLEAR);
    }
};

} // namespace com
