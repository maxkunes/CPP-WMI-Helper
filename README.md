# Windows WMI Modern C++17 Helper
## This single header helper provides a modern interface to WMI.

## Usecase:
### This interface was written for a project I am working on. This is really what I would concider to be a first draft. Bugs exist and documentation is non existent.

## Example:
```
wmi_helper<wmi_any32> wmi_helper_;
std::uint64_t bank_handle_ = 0;
std::uint64_t speed_handle_ = 0;
std::uint64_t capacity_handle_ = 0;
std::uint64_t locator_handle_ = 0;

bool init()
{
    // tell helper what class we want data from. aka what table
    if (!wmi_helper_.init(L"Win32_PhysicalMemory"))
    {
        wmi_helper_.cleanup();
        return false;
    }

    // returns optional handle for var aka column of table
    auto bound_var = wmi_helper_.bind_var(L"BankLabel");

    // if no handle was returned, cleanup and return out of method
    if (!bound_var) {
        wmi_helper_.cleanup();
        return false;
    }

    // grab handle value from optional for later
    bank_handle_ = bound_var.value();

    bound_var = wmi_helper_.bind_var(L"Speed");

    if (!bound_var) {
        wmi_helper_.cleanup();
        return false;
    }

    speed_handle_ = bound_var.value();

    bound_var = wmi_helper_.bind_var(L"Capacity");

    if (!bound_var) {
        wmi_helper_.cleanup();
        return false;
    }

    capacity_handle_ = bound_var.value();

    bound_var = wmi_helper_.bind_var(L"DeviceLocator");

    if (!bound_var) {
        wmi_helper_.cleanup();
        return false;
    }

    locator_handle_ = bound_var.value();

    // can be bound to class member functions using std::bind and placeholders
    wmi_helper_.bind_callback(wmi_callback);

    // starts the update thread to start querying wmi
    wmi_helper_.start();

    return true;
}

// Callback from another thread. Not running in main thread/init thread.
void wmi_callback(wmi_wrapper_results<wmi_any32>& results, wmi_wrapper_results<wmi_any32>& prev_results)
{
    if (!results.empty() && !prev_results.empty())
    {
        if (results.count(bank_handle_) && results.count(speed_handle_) && results.count(capacity_handle_) && results.count(locator_handle_))
        {

            auto& bank_results = results[bank_handle_];
            auto& speed_results = results[speed_handle_];
            auto& capacity_results = prev_results[capacity_handle_];
            auto& locator_results = prev_results[locator_handle_];

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

```