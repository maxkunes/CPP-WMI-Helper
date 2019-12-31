#pragma once
#include <windows.h>
#include <WbemCli.h>
#include <thread>
#include <map>
#include <optional>
#include <functional>
#include <atomic>
#include <future>
#include <utility>
#include <vector>
#include <mutex>
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "Propsys.lib")

#include "fmt/format.h"

inline std::uint64_t get_current_time()
{
    return std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::system_clock::now().time_since_epoch()).count();
}

template<std::size_t MaxSize = 32>
struct wmi_any
{
    CIMTYPE type;
    char reserved[MaxSize];
    std::wstring str;

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
        return str;
    }

    [[nodiscard]] std::string get_string() const
    {
        return std::string(str.begin(), str.end());
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
struct wmi_wrapper_class_result
{
    wmi_wrapper_result_map<AnySize> result;
    wmi_wrapper_result_map<AnySize> prev_result;
};

using wmi_wrapper_32_class_result = wmi_wrapper_class_result<32>;

template<std::size_t AnySize>
using wmi_helper_callback = std::function<void(const wmi_helper_config&, const wmi_wrapper_class_result<AnySize>&)>;

template<std::size_t AnySize>
using wmi_wrapper_vector_result = std::vector<wmi_wrapper_class_result<AnySize>>;

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

    void init(const wmi_helper_config& config)
    {
        unsigned long hr = S_OK;
        long    l_id_ = 0;

        config_ = config;

        if (FAILED(hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED)))
        {
            cleanup();
            throw std::exception(fmt::format("CoInitializeEx failed with error code {0:#x}.", hr).c_str());
        }

        if (FAILED(hr = CoInitializeSecurity(
            nullptr,
            -1,
            nullptr,
            nullptr,
            RPC_C_AUTHN_LEVEL_NONE,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            NULL, EOAC_NONE, nullptr)))
        {
            if (static_cast<HRESULT>(hr) != RPC_E_TOO_LATE) {
                cleanup();
                throw std::exception(fmt::format("CoInitializeSecurity failed with error code {0:#x}.", hr).c_str());
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
            throw std::exception(fmt::format("CoCreateInstance failed with error code {0:#x}.", hr).c_str());
        }

        // Connect to the desired namespace.
        bstr_name_space_ = SysAllocString(config_.server().c_str());
    	
        if (nullptr == bstr_name_space_)
        {
            cleanup();
            throw std::exception(fmt::format("SysAllocString failed.").c_str());
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
            throw std::exception(fmt::format("ConnectServer failed with error code {0:#x}.", hr).c_str());
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
            throw std::exception(fmt::format("CoCreateInstance failed with error code {0:#x}.", hr).c_str());
        }

        if (FAILED(hr = p_refresher_->QueryInterface(
            IID_IWbemConfigureRefresher,
            reinterpret_cast<void**>(&p_config_))))
        {
            cleanup();
            throw std::exception(fmt::format("QueryInterface failed with error code {0:#x}.", hr).c_str());
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
            throw std::exception(fmt::format("AddEnum failed with error code {0:#x}.", hr).c_str());
        }

        p_config_->Release();
        p_config_ = nullptr;
    }

    void stop_query()
    {
        if (querying_) {

            querying_ = false;

            // query_async will set the signal to true when terminating. possible race condition?
            while (querying_)
            {
                std::this_thread::sleep_for(std::chrono::milliseconds(500));
            }

            querying_ = false;
        }
    }
	
    void cleanup()
    {
    	// TODO: better thread termination signal system
        stop_query();


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
	
    wmi_wrapper_vector_result<AnySize> query()
    {
        if(querying_)
        {
            throw std::exception("Cannot start query while another one is already running!");
        }
    	
        if(config_.fire_count() == wmi_helper_config::infinite
            && config_.fire_time() == wmi_helper_config::infinite)
        {
            throw std::exception("WmiHelper::query() (non async) cannot be called with an infinite fire_count and fire_time as it would never complete!");
        }
    	
        return query_internal(false, true, nullptr, bound_vars_, config_, get_current_time()).value();
    }

    std::future<void> query_async(const wmi_helper_callback<AnySize>& callback)
    {
        auto current_time = get_current_time();
        auto config = config_;
        auto vars = bound_vars_;
    	
        return std::async(std::launch::async, [this, callback, current_time, vars, config]()
        {
	        query_internal(true, false, callback, vars, config, current_time);
        });
    }

    std::future <wmi_wrapper_vector_result<AnySize>> query_async_return(const wmi_helper_callback<AnySize>& callback)
    {
        auto current_time = get_current_time();
        auto config = config_;
        auto vars = bound_vars_;

        return std::async(std::launch::async, [this, callback, current_time, vars, config]()
        {
			auto opt = query_internal(true, true, callback, vars, config, current_time);
            return opt.value(); // no way to ever have no value
        });
    }
	

private:
	
    std::optional<wmi_wrapper_vector_result<AnySize>> query_internal(const bool async, const bool return_data, const wmi_helper_callback<AnySize> callback, const std::unordered_map<std::uint64_t, std::wstring> bound_vars, const wmi_helper_config& config, const std::uint64_t start_time)
    {
        auto fire_count = 0;
        wmi_wrapper_vector_result<AnySize> ret_value;

        wmi_wrapper_result_map<AnySize> results_;
        wmi_wrapper_result_map<AnySize> prev_results_;

        querying_ = true;
    	
        while (true)
        {
            if (async)
            {
                if (!querying_)
                {
                    querying_ = true;

                    if (return_data)
                        return ret_value;
                	
                    return std::nullopt;
                }
            }
        	
            results_.clear();

            HRESULT hr = S_OK;
        	
            const auto num_rows = refresh_data();

            if (!ap_enum_access_ || !ap_enum_access_[0])
                continue;

            for (auto& bound_var : bound_vars)
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

                	if(var_type == 8)
                	{
                        // wstring

                        if (FAILED(ap_enum_access_[i]->ReadPropertyValue(var_handle, 0, &read_bytes, any.data<byte*>())))
                        {
                            std::wstring str;
                            str.resize(read_bytes);

                            if (FAILED(ap_enum_access_[i]->ReadPropertyValue(var_handle, read_bytes, &read_bytes, reinterpret_cast<byte*>(str.data()))))
                            {
                                continue;
                            }

                            any.str = str;
                        }
                		
                	}
                    else if (FAILED(ap_enum_access_[i]->ReadPropertyValue(var_handle, sizeof(any), &read_bytes, any.data<byte*>())))
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

        	if(async && !return_data)
        	{
				callback(config, { results_, prev_results_ });
        	}
            else {  // NOLINT(readability-misleading-indentation)
				ret_value.push_back({ results_, prev_results_ });
            }
        	
            prev_results_ = results_;

            fire_count++;

            if (config.fire_count() != wmi_helper_config::infinite) {
                if (fire_count == config.fire_count())
                {
                    querying_ = false;
                	
                    if(async && !return_data)
                    {
                        return std::nullopt;
                    }
                	
                    return ret_value;
                }
            }

            if (config.fire_time() != wmi_helper_config::infinite) {
                if (get_current_time() >= start_time + config.fire_time())
                {
                    querying_ = false;
                    if (async && !return_data)
                    {
                        return std::nullopt;
                    }
                	
                    return ret_value;
                }
            }

            std::this_thread::sleep_for(std::chrono::milliseconds(1000 / config.updates_per_second()));
        }

        querying_ = false;
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

    std::unordered_map<std::uint64_t, std::wstring> bound_vars_;
	
    std::int32_t updates_per_second_ = 1;
	
    std::atomic_bool querying_ = false;

    wmi_helper_config config_;
};

using wmi_helper_32 = wmi_helper <32>;
using wmi_helper_32_results = wmi_wrapper_result_map<32>;