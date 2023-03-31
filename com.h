#pragma once

#include <Windows.h>
#include <string>
#include <cassert>

namespace com
{

class ComPort
{
public:
    typedef HANDLE port_t;

private:
    port_t m_port{ nullptr };
    unsigned int m_com_index{ 0 };

public:
    ComPort(unsigned int com_index) : m_com_index(com_index) {}
    ComPort(unsigned int com_index, unsigned int baud_rate, unsigned int parity,
        unsigned int data_bits, unsigned int stop_bits) : m_com_index(com_index) 
    {
        set_baud_rate(baud_rate);
        set_parity(parity);
        set_data_bits(data_bits);
        set_stop_bits(stop_bits);
    }

    ~ComPort()
    {
        flush();
        close();
    }

    void set_baud_rate(unsigned int baud_rate)
    {
        DCB dcbSerialParams = { 0 };
        BOOL Status;
        dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
        Status = GetCommState(m_port, &dcbSerialParams);

        dcbSerialParams.BaudRate = baud_rate;
        Status = SetCommState(m_port, &dcbSerialParams);
    }

    void set_parity(unsigned int parity)
    {
        DCB dcbSerialParams = { 0 };
        BOOL Status;
        dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
        Status = GetCommState(m_port, &dcbSerialParams);

        dcbSerialParams.Parity = parity;
        Status = SetCommState(m_port, &dcbSerialParams);

    }

    void set_data_bits(unsigned int bits)
    {
        DCB dcbSerialParams = { 0 };
        BOOL Status;
        dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
        Status = GetCommState(m_port, &dcbSerialParams);

        dcbSerialParams.ByteSize = bits;
        Status = SetCommState(m_port, &dcbSerialParams);
    }

    void set_stop_bits(unsigned int bits)
    {
        DCB dcbSerialParams = { 0 };
        BOOL Status;
        dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
        Status = GetCommState(m_port, &dcbSerialParams);
        
        dcbSerialParams.StopBits = bits;
        Status = SetCommState(m_port, &dcbSerialParams);
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
            m_port = port;
        }
    }

    void close()
    {
        CloseHandle(m_port);
    }

    bool write(const std::string& data)
    {
        DWORD bytesWritten;
        bool status = WriteFile(
            m_port,
            data.c_str(),
            data.length(),
            &bytesWritten,
            NULL);

        return status;
    }

    bool read(std::string& data, unsigned int num_bytes)
    {
        DWORD dwEventMask;
        DWORD NoBytesRead;
        BOOL status = WaitCommEvent(m_port, &dwEventMask, NULL);
        if (!status)
        {
            return false;
        }

        char* raw_data = nullptr;
        bool status = ReadFile(m_port, raw_data, num_bytes, &NoBytesRead, NULL);
        data = std::string{ raw_data };
        
        return status;
    }

    void flush()
    {
        PurgeComm(m_port, PURGE_RXCLEAR);
        PurgeComm(m_port, PURGE_TXCLEAR);
    }
};

} // namespace com
