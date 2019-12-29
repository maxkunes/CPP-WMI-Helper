#include <iostream>
#include "WmiHelper.hpp"

// WmiHelper.hpp uses std::optional and structured bindings. Either compile with a supported c++ version or remove the uses of optional and structured bindings. It is like one function.

wmi_helper_32 helper;
std::uint64_t bank_handle = 0;
std::uint64_t speed_handle = 0;
std::uint64_t capacity_handle = 0;
std::uint64_t locator_handle = 0;

// Executes in its own thread. Use mutexes or whatever to synchronize between threads for data passing.
void wmi_callback(const wmi_helper_config& config, const wmi_wrapper_32_result& wmi_result)
{
    auto& results = wmi_result.result;
    auto prev_results = wmi_result.prev_result;
	
    if (!results.empty())
    {
        if (results.count(bank_handle) && results.count(speed_handle) && results.count(capacity_handle) && results.count(locator_handle))
        {

            auto& bank_results = results.at(bank_handle);
            auto& speed_results = results.at(speed_handle);
            auto& capacity_results = results.at(capacity_handle);
            auto& locator_results = results.at(locator_handle);

            for (auto i = 0; i < bank_results.size(); i++)
            {
                // not exactly safe. 32 bytes may be too small for the string. check the wmi_any32 class for more details.
                // we are assuming each column has the same amount of rows. This is not always the case if a property read fails due to something like insufficient buffer size.
                const auto bank = bank_results[i].get_string();
                const auto speed = speed_results[i].get<std::uint32_t>();
                const auto capacity = capacity_results[i].get<std::uint64_t>();
                const auto locator = locator_results[i].get_string();

                // cout is not thread safe as far as I know. purely used as an example. functions for me, might not for you. send a message back to the main thread using something like mutexes.
                std::cout << bank << " at " << locator << " has a speed of " << speed << "mhz and a capacity of " << capacity << std::endl;
            }

        }

    }
}


void main()
{
	const wmi_helper_config config(
        L"Win32_PhysicalMemory",
        5,
        -1,
        1);
	
    // tell helper what class we want data from. aka what table
    if (!helper.init(config))
    {
        helper.cleanup();
        return;
    }
	
    bank_handle = helper.capture_var(L"BankLabel");
    speed_handle = helper.capture_var(L"Speed");
    capacity_handle = helper.capture_var(L"Capacity");
    locator_handle = helper.capture_var(L"DeviceLocator");
	
	const auto& sync_result_opt = helper.query();

    if(sync_result_opt.has_value())
    {
        const auto& sync_result = sync_result_opt.value();

        for (auto& [results, prev_results] : sync_result) {

            if (results.count(bank_handle) && results.count(speed_handle) && results.count(capacity_handle) && results.count(locator_handle))
            {

                auto& bank_results = results.at(bank_handle);
                auto& speed_results = results.at(speed_handle);
                auto& capacity_results = results.at(capacity_handle);
                auto& locator_results = results.at(locator_handle);

                for (auto i = 0; i < bank_results.size(); i++)
                {
                    // not exactly safe. 32 bytes may be too small for the string. check the wmi_any32 class for more details.
                    // we are assuming each column has the same amount of rows. This is not always the case if a property read fails due to something like insufficient buffer size.
                    const auto bank = bank_results[i].get_string();
                    const auto speed = speed_results[i].get<std::uint32_t>();
                    const auto capacity = capacity_results[i].get<std::uint64_t>();
                    const auto locator = locator_results[i].get_string();

                    // cout is not thread safe as far as I know. purely used as an example. functions for me, might not for you. send a message back to the main thread using something like mutexes.
                    std::cout << bank << " at " << locator << " has a speed of " << speed << "mhz and a capacity of " << capacity << std::endl;
                }

            }
        }
    }
	

    helper.query_async(wmi_callback);
	
    std::this_thread::sleep_for(std::chrono::milliseconds(5000));
	
    return;
}
