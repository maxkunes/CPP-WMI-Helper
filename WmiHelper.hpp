#pragma once
#include <windows.h>
#include <WbemCli.h>
#include <iostream>
#include <thread>
#include <map>
#include <optional>
#include <functional>
#include <atomic>
#include <vector>
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "Propsys.lib")

template<std::size_t MaxSize = 32>
struct wmi_any
{
    CIMTYPE type;
    char reserved[MaxSize];

    void* data()
    {
        return static_cast<void*>(reserved);
    }

    template<typename T>
    T& get()
    {
        static_assert(sizeof(T) <= MaxSize, "sizeof(T) too large to fit in WmiAny buffer. Please increase MaxSize template to fix this issue.");
        return *reinterpret_cast<T*>(data());
    }
};

using wmi_any32 = wmi_any<32>;

template<typename WmiAnyType>
using wmi_wrapper_results = std::map< std::uint64_t, std::vector<WmiAnyType>>;

template<typename WmiAnyType>
using wmi_helper_callback = std::function<void(wmi_wrapper_results<WmiAnyType>&, wmi_wrapper_results<WmiAnyType>&)>;

template<typename WmiAnyType>
class wmi_helper
{
public:

    wmi_helper() = default;

    ~wmi_helper()
    {
        cleanup();
    }

    bool init(const std::wstring& class_name, const std::int32_t updates_per_second = 1)
    {
        HRESULT hr = S_OK;
        long    l_id_ = 0;
        updates_per_second_ = updates_per_second;

        if (FAILED(hr = CoInitializeEx(NULL, COINIT_MULTITHREADED)))
        {
            cleanup();
            return false;
        }

        if (FAILED(hr = CoInitializeSecurity(
            NULL,
            -1,
            NULL,
            NULL,
            RPC_C_AUTHN_LEVEL_NONE,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            NULL, EOAC_NONE, nullptr)))
        {
            cleanup();
            return false;
        }

        if (FAILED(hr = CoCreateInstance(
            CLSID_WbemLocator,
            NULL,
            CLSCTX_INPROC_SERVER,
            IID_IWbemLocator,
            reinterpret_cast<void**>(&p_wbem_locator_))))
        {
            cleanup();
            return false;
        }

        // Connect to the desired namespace.
        bstr_name_space_ = SysAllocString(L"\\\\.\\root\\cimv2");
        if (nullptr == bstr_name_space_)
        {
            cleanup();
            return false;
        }
        if (FAILED(hr = p_wbem_locator_->ConnectServer(
            bstr_name_space_,
            NULL, // User name
            NULL, // Password
            NULL, // Locale
            0L,   // Security flags
            NULL, // Authority
            NULL, // Wbem context
            &p_name_space_)))
        {

            cleanup();
            return false;
        }

        p_wbem_locator_->Release();
        p_wbem_locator_ = nullptr;
        SysFreeString(bstr_name_space_);
        bstr_name_space_ = nullptr;

        if (FAILED(hr = CoCreateInstance(
            CLSID_WbemRefresher,
            NULL,
            CLSCTX_INPROC_SERVER,
            IID_IWbemRefresher,
            reinterpret_cast<void**>(&p_refresher_))))
        {
            cleanup();
            return false;
        }

        if (FAILED(hr = p_refresher_->QueryInterface(
            IID_IWbemConfigureRefresher,
            reinterpret_cast<void**>(&p_config_))))
        {
            cleanup();
            return false;
        }

        // Add an enumerator to the refresher.
        if (FAILED(hr = p_config_->AddEnum(
            p_name_space_,
            class_name.c_str(),
            0,
            NULL,
            &p_enum_,
            &l_id_)))
        {
            cleanup();
            return false;
        }

        p_config_->Release();
        p_config_ = nullptr;


        return true;
    }

    void cleanup()
    {
        if (thread_running && update_thread_) {

            thread_close_signal_ = true;
            update_thread_->join();
        	
            // update_thread will set the signal to false when ready to terminate.
            while (thread_close_signal_)
            {
                std::this_thread::sleep_for(std::chrono::milliseconds(500));
            }
        }


        if (p_wbem_locator_) {
            p_wbem_locator_->Release();
            p_wbem_locator_ = nullptr;
        }

        if (p_name_space_)
        {
            p_name_space_->Release();
            p_name_space_ = nullptr;
        }

        if (bstr_name_space_) {
            SysFreeString(bstr_name_space_);
            bstr_name_space_ = nullptr;
        }

        if (p_refresher_)
        {
            p_refresher_->Release();
            p_refresher_ = nullptr;
        }

        if (p_config_)
        {
            p_config_->Release();
            p_config_ = nullptr;
        }

        if (p_enum_)
        {
            p_enum_->Release();
            p_enum_ = nullptr;
        }

        if (ap_enum_access_)
        {
            delete[] ap_enum_access_;
            ap_enum_access_ = nullptr;
        }

        CoUninitialize();
    }

    void bind_callback(const wmi_helper_callback<WmiAnyType>& callback)
    {
        callback_ = callback;
    }

    std::optional<std::uint64_t> bind_var(const std::wstring& var_name)
    {
        const auto var_hash = std::hash<std::wstring>{}(var_name);

        if (bound_vars_.count(var_hash))
            return std::nullopt;

        bound_vars_[var_hash] = var_name;

        return var_hash;
    }

    void start()
    {
        update_thread_ = std::make_unique<std::thread>(&wmi_helper::update_thread, this);
        thread_running = true;
    }
	
    void update()
    {
        HRESULT hr = S_OK;
        DWORD   dw_num_returned = 0;

        results_.clear();

        if (FAILED(hr = p_refresher_->Refresh(0L)))
        {
            return;
        }

        if (ap_enum_access_length_ > 0)
        {
            SecureZeroMemory(ap_enum_access_,
                ap_enum_access_length_ * sizeof(IWbemObjectAccess*));
        }

        hr = p_enum_->GetObjects(0L,
            ap_enum_access_length_,
            ap_enum_access_,
            &dw_num_returned);
        // If the buffer was not big enough,
        // allocate a bigger buffer and retry.
        if (hr == WBEM_E_BUFFER_TOO_SMALL
            && dw_num_returned > ap_enum_access_length_)
        {
            ap_enum_access_ = new IWbemObjectAccess * [dw_num_returned];
            if (nullptr == ap_enum_access_)
            {
                hr = E_OUTOFMEMORY;
                return;
            }

            SecureZeroMemory(ap_enum_access_,
                dw_num_returned * sizeof(IWbemObjectAccess*));
            ap_enum_access_length_ = dw_num_returned;

            if (FAILED(hr = p_enum_->GetObjects(0L,
                ap_enum_access_length_,
                ap_enum_access_,
                &dw_num_returned)))
            {
                return;
            }
        }

        if (!ap_enum_access_ || !ap_enum_access_[0])
            return;

        for (auto& bound_var : bound_vars_)
        {
            auto& [hash, name] = bound_var;

            CIMTYPE var_type;
            long var_handle;

            if (FAILED(hr = ap_enum_access_[0]->GetPropertyHandle(
                name.c_str(),
                &var_type,
                &var_handle)))
            {
                continue;
            }

            for (auto i = 0; i < dw_num_returned; i++) {

                WmiAnyType any{};
                any.type = var_type;

                long read_bytes = 0x0;


                if (FAILED(ap_enum_access_[i]->ReadPropertyValue(var_handle, sizeof(any), &read_bytes, reinterpret_cast<byte*>(any.data()))))
                {
                    continue;
                }

                results_[hash].push_back(any);
            }

        }

        for (auto i = 0; i < dw_num_returned; i++) {
            ap_enum_access_[i]->Release();
            ap_enum_access_[i] = nullptr;
        }

        callback_(results_, prev_results_);
        prev_results_ = results_;

        return;
    }

    void update_thread()
    {
        while (true)
        {
            if (thread_close_signal_)
            {
                thread_close_signal_ = false;
                thread_running = false;
                return;
            }

            update();
            std::this_thread::sleep_for(std::chrono::milliseconds(1000 / updates_per_second_));
        }
    }
private:
    // Used by init method
    IWbemConfigureRefresher* p_config_ = nullptr;
    IWbemServices* p_name_space_ = nullptr;
    IWbemLocator* p_wbem_locator_ = nullptr;
    BSTR bstr_name_space_ = nullptr;

    // Used by update_thread method
    IWbemObjectAccess** ap_enum_access_ = nullptr;
    DWORD ap_enum_access_length_ = 0;
    IWbemHiPerfEnum* p_enum_ = nullptr;
    IWbemRefresher* p_refresher_ = nullptr;

    std::unordered_map<std::uint64_t, std::wstring> bound_vars_;
    wmi_wrapper_results<WmiAnyType> results_;
    wmi_wrapper_results<WmiAnyType> prev_results_;
    wmi_helper_callback<WmiAnyType> callback_;

    std::unique_ptr<std::thread> update_thread_;
    std::int32_t updates_per_second_ = 1;

    std::atomic_bool thread_close_signal_ = false;
    std::atomic_bool thread_running = false;
};