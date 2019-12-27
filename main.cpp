#include <iostream>
#include "WmiHelper.hpp"

// WmiHelper.hpp uses std::optional. Either compile with a supported c++ version or remove the uses of optional. It is like one function.

std::uint64_t name_handle_ = 0;
std::uint64_t usage_handle_ = 0;
std::uint64_t time_handle_ = 0;

// Executes in its own thread. Use mutexes or whatever to synchronize between threads for data passing.
void wmi_callback(wmi_wrapper_results<wmi_any32>& results, wmi_wrapper_results<wmi_any32>& prev_results)
{
    if (!results.empty() && !prev_results.empty())
    {
        if (results.count(time_handle_) && results.count(usage_handle_) && results.count(name_handle_))
        {
            if (prev_results.count(time_handle_) && prev_results.count(usage_handle_))
            {
                auto& name_results = results[name_handle_];
                auto& time_results = results[time_handle_];
                auto& usage_results = results[usage_handle_];
                auto& prev_time_results = prev_results[time_handle_];
                auto& prev_usage_results = prev_results[usage_handle_];

                for (auto i = 0; i < time_results.size(); i++)
                {
                    auto thread_name_wide = std::wstring(static_cast<const wchar_t*>(name_results[i].data()));

                	// Please forgive me...
                    std::string thread_name = std::string(thread_name_wide.begin(), thread_name_wide.end());
                	
                    const double new_usage = usage_results[i].get<std::uint64_t>();
                    const double old_usage = prev_usage_results[i].get<std::uint64_t>();
                    const double new_time = time_results[i].get<std::uint64_t>();
                    const double old_time = prev_time_results[i].get<std::uint64_t>();

                    const auto percent_processor_time = (1 - ((new_usage - old_usage) / (new_time - old_time))) * 100.0;

                    std::cout << "Thread #" << thread_name << " has " << percent_processor_time << "% usage.\r\n";
                }

            }
        }
    }
}


void main()
{
    wmi_helper<wmi_any32> wmi_helper_;


    if (!wmi_helper_.init(L"Win32_PerfRawData_PerfOS_Processor", 2 /* updates_per_second */))
    {
        wmi_helper_.cleanup();
        return;
    }

    auto bound_var = wmi_helper_.bind_var(L"PercentProcessorTime");

    if (!bound_var) {
        wmi_helper_.cleanup();
        return;
    }

    usage_handle_ = bound_var.value();

    bound_var = wmi_helper_.bind_var(L"TimeStamp_Sys100NS");

    if (!bound_var) {
        wmi_helper_.cleanup();
        return;
    }

    time_handle_ = bound_var.value();

    bound_var = wmi_helper_.bind_var(L"Name");

    if (!bound_var) {
        wmi_helper_.cleanup();
        return;
    }

    name_handle_ = bound_var.value();

    wmi_helper_.bind_callback(wmi_callback);

    wmi_helper_.start();

    // Retrieve data for 10 seconds.
    std::this_thread::sleep_for(std::chrono::seconds(2));

    // automatically called by destructor, available if you want to manually call.
    wmi_helper_.cleanup();
    return;
}
