#pragma once
#include <windows.h>
#include <WbemCli.h>
#include <iostream>
#include <thread>
#include <map>
#include <optional>
#include <functional>
#include <atomic>
#include <utility>
#include <vector>
#include <mutex>
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "Propsys.lib")

inline std::uint64_t get_current_time()
{
    return std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::system_clock::now().time_since_epoch()).count();
}

template<std::size_t MaxSize = 32>
struct wmi_any
{
    CIMTYPE type;
    char reserved[MaxSize];

	template<typename T>
    [[nodiscard]] T data() const
    {
	    // ReSharper disable once CppCStyleCast
	    return (T)(reserved);
    }

    template<typename T>
    [[nodiscard]] T& get() const
    {
        static_assert(sizeof(T) <= MaxSize, "sizeof(T) too large to fit in WmiAny buffer. Please increase MaxSize template to fix this issue.");
        return *reinterpret_cast<T*>(data<void*>());
    }

    [[nodiscard]] std::wstring get_wide_string() const
    {
	    const auto ptr = data<const wchar_t*>();
		
        return std::wstring(ptr);
    }

    [[nodiscard]] std::string get_string() const
    {
        auto wide_string = get_wide_string();
        return std::string(wide_string.begin(), wide_string.end());
    }

};



class wmi_helper_config
{
public:
	
    wmi_helper_config(std::wstring class_name, std::int32_t fire_count = infinite, std::int32_t fire_time = infinite, std::int32_t updates_per_second = 2) : class_name_(std::move(class_name)), fire_count_(fire_count), fire_time_(fire_time), updates_per_second_(updates_per_second)
    {

    }

    wmi_helper_config() = default;

    [[nodiscard]] std::wstring& server()
    {
        return server_;
    }

    [[nodiscard]] std::wstring& username()
    {
        return username_;
    }

    [[nodiscard]] std::wstring& password()
    {
        return password_;
    }

	
    [[nodiscard]] const std::wstring& class_name() const
    {
        return class_name_;
    }
	
    [[nodiscard]] const std::int32_t& fire_count() const
    {
        return fire_count_;
    }

    [[nodiscard]] const std::int32_t& fire_time() const
    {
        return fire_time_;
    }

    [[nodiscard]] const std::int32_t& updates_per_second() const
    {
        return updates_per_second_;
    }

    const static std::int32_t infinite = -1;
private:
    std::wstring class_name_;
    std::int32_t fire_count_ = -1; // -1 for infinity
    std::int32_t fire_time_ = 5000; // -1 for infinity
    std::int32_t updates_per_second_ = 2; // times wmi is queried per second

    std::wstring server_ = L"\\\\.\\root\\cimv2";
    std::wstring username_;
    std::wstring password_;

};

using wmi_var_handle = std::uint64_t;

template<std::size_t AnySize>
using wmi_wrapper_result_map = std::map< wmi_var_handle, std::vector<wmi_any<AnySize>>>;

template<std::size_t AnySize>
struct wmi_wrapper_result
{
    wmi_wrapper_result_map<AnySize> result;
    wmi_wrapper_result_map<AnySize> prev_result;
};

using wmi_wrapper_32_result = wmi_wrapper_result<32>;

template<std::size_t AnySize>
using wmi_helper_callback = std::function<void(const wmi_helper_config&, const wmi_wrapper_result<AnySize>&)>;

template<std::size_t AnySize>
using wmi_wrapper_sync_result = std::vector<wmi_wrapper_result<AnySize>>;

template<std::size_t AnySize>
class wmi_helper
{
public:

    wmi_helper()
    = default;

    ~wmi_helper()
    {
        cleanup();
    }

    bool init(const wmi_helper_config& config)
    {
        HRESULT hr = S_OK;
        long    l_id_ = 0;

        config_ = config;

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
            if (hr != RPC_E_TOO_LATE) {
                cleanup();
                return false;
            }
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
        bstr_name_space_ = SysAllocString(config_.server().c_str());
    	
        if (nullptr == bstr_name_space_)
        {
            cleanup();
            return false;
        }

        BSTR username = nullptr;

        if (!config_.username().empty())
            username = SysAllocString(config_.username().c_str());

        BSTR password = nullptr;

        if (!config_.password().empty())
            password = SysAllocString(config_.password().c_str());
    	
        if (FAILED(hr = p_wbem_locator_->ConnectServer(
            bstr_name_space_,
            username, // User name
            password, // Password
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

    	if(username)
    	{
            SysFreeString(username);
            username = nullptr;
    	}

    	if(password)
    	{
            SysFreeString(password);
            password = nullptr;
    	}
             
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
            config_.class_name().c_str(),
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

        if (update_thread_->joinable())
            update_thread_->join();


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

    wmi_var_handle capture_var(const std::wstring& var_name)
    {
        const auto var_hash = std::hash<std::wstring>{}(var_name);

        bound_vars_[var_hash] = var_name;

        return var_hash;
    }

    std::uint32_t refresh_data()
    {
        HRESULT hr = S_OK;
        ULONG dw_num_returned;

        if (FAILED(hr = p_refresher_->Refresh(0L)))
        {
            return 0;
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
                return 0;
            }

            SecureZeroMemory(ap_enum_access_,
                dw_num_returned * sizeof(IWbemObjectAccess*));
            ap_enum_access_length_ = dw_num_returned;

            if (FAILED(hr = p_enum_->GetObjects(0L,
                ap_enum_access_length_,
                ap_enum_access_,
                &dw_num_returned)))
            {
                return 0;
            }
        }

        return dw_num_returned;
    }
	
    std::optional<wmi_wrapper_sync_result<AnySize>> query()
    {
        if(config_.fire_count() == wmi_helper_config::infinite
            && config_.fire_time() == wmi_helper_config::infinite)
        {
            throw std::exception("WmiHelper::query() (non async) cannot be called with an infinite fire_count and fire_time as it would never complete!");
        }
    	
        return query_internal(config_, get_current_time());
    }

    void query_async(const wmi_helper_callback<AnySize>& callback)
    {
        update_thread_ = std::make_unique<std::thread>(&wmi_helper::query_async_internal, this, callback, config_, get_current_time());
        thread_running = true;
    }

private:
	
    void query_async_internal(const wmi_helper_callback<AnySize>& callback, const wmi_helper_config config, const std::uint64_t start_time)
    {
        auto fire_count = 0;
        wmi_wrapper_result_map<AnySize> results_;
        wmi_wrapper_result_map<AnySize> prev_results_;
    	
        while (true) {
            results_.clear();
            HRESULT hr = S_OK;

            if (thread_close_signal_)
            {
                thread_close_signal_ = false;
                thread_running = false;
                return;
            }

            const auto num_rows = refresh_data();

            if (!ap_enum_access_ || !ap_enum_access_[0])
                return;


            auto bound_vars_lock = std::scoped_lock(bound_vars_mutex_);

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

                for (auto i = 0; i < num_rows; i++) {

                    wmi_any<AnySize> any{};
                    any.type = var_type;

                    long read_bytes = 0x0;


                    if (FAILED(ap_enum_access_[i]->ReadPropertyValue(var_handle, sizeof(any), &read_bytes, any.data<byte*>())))
                    {
                        continue;
                    }


                    results_[hash].push_back(any);
                }

            }

            for (auto i = 0; i < num_rows; i++) {
                ap_enum_access_[i]->Release();
                ap_enum_access_[i] = nullptr;
            }



            callback(config, { results_, prev_results_ });
            prev_results_ = results_;

            fire_count++;

            if (config.fire_count() != wmi_helper_config::infinite) {
                if (fire_count == config.fire_count())
                {
                    thread_running = false;
                    return;
                }
            }

            if (config.fire_time() != wmi_helper_config::infinite) {
                if (get_current_time() >= start_time + config.fire_time())
                {
                    thread_running = false;
                    return;
                }
            }

            std::this_thread::sleep_for(std::chrono::milliseconds(1000 / config.updates_per_second()));
        }
        return;
    }

    std::optional<wmi_wrapper_sync_result<AnySize>> query_internal(const wmi_helper_config& config, const std::uint64_t start_time)
    {
        auto fire_count = 0;
        wmi_wrapper_sync_result<AnySize> ret_value;

        wmi_wrapper_result_map<AnySize> results_;
        wmi_wrapper_result_map<AnySize> prev_results_;
    	
        while (true)
        {
            results_.clear();

            HRESULT hr = S_OK;

            const auto num_rows = refresh_data();

            if (!ap_enum_access_ || !ap_enum_access_[0])
                continue;

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

                for (auto i = 0; i < num_rows; i++) {

                    wmi_any<AnySize> any{};
                    any.type = var_type;

                    long read_bytes = 0x0;


                    if (FAILED(ap_enum_access_[i]->ReadPropertyValue(var_handle, sizeof(any), &read_bytes, any.data<byte*>())))
                    {
                        continue;
                    }

                    results_[hash].push_back(any);
                }

            }

            for (auto i = 0; i < num_rows; i++) {
                ap_enum_access_[i]->Release();
                ap_enum_access_[i] = nullptr;
            }

            ret_value.push_back({ results_, prev_results_ });
            prev_results_ = results_;

            fire_count++;

            if (config.fire_count() != wmi_helper_config::infinite) {
                if (fire_count == config.fire_count())
                {
                    return ret_value;
                }
            }

            if (config.fire_time() != wmi_helper_config::infinite) {
                if (get_current_time() >= start_time + config.fire_time())
                {
                    return ret_value;
                }
            }

            std::this_thread::sleep_for(std::chrono::milliseconds(1000 / config.updates_per_second()));
        }

        return std::nullopt;
    }

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

    std::mutex bound_vars_mutex_;
    std::unordered_map<std::uint64_t, std::wstring> bound_vars_;

    std::unique_ptr<std::thread> update_thread_;
    std::int32_t updates_per_second_ = 1;

    std::atomic_bool thread_close_signal_ = false;
    std::atomic_bool thread_running = false;

    wmi_helper_config config_;
};

using wmi_helper_32 = wmi_helper <32>;
using wmi_helper_32_results = wmi_wrapper_result_map<32>;