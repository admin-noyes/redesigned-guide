# Time Complexity Analysis of URL Validator Excel Add-In

## Overview
The URL Validator Add-In processes a list of URLs by performing HTTP HEAD requests to check their validity. The updated version incorporates multi-threaded checking using PowerShell for parallel execution, significantly improving performance for large datasets. This analysis examines the time complexity of both the original single-threaded and the new parallel versions.

## Code Analysis

### Original Single-Threaded Version
- **Main Loop**: Iterates sequentially through each URL (O(n) where n = number of URLs).
- **URL Checking**: Calls `CheckURLStatus` with up to 3 retries per URL.
- **HTTP Requests**: Uses `WinHttp.WinHttpRequest.5.1` with timeouts; each request is blocking.
- **Time Complexity**: O(n) - Network I/O dominates, with each URL taking 1-5 seconds on average.

### Updated Parallel Version
- **URL Collection**: Gathers all URLs into an array (O(n)).
- **PowerShell Integration**: Writes URLs to a temp file, invokes a PowerShell script for parallel processing, reads results back.
- **Parallel Checking**: Uses PowerShell jobs to check up to 10 URLs concurrently (configurable).
- **Key Changes**:
  - Eliminates sequential blocking; jobs run in parallel.
  - Overhead: File I/O for input/output, job management.
  - Progress tracking removed (parallel nature makes per-URL progress complex).
- **Time Complexity**: O(n/k + overhead) where k = number of parallel threads (default 10). The network-bound operations are parallelized, reducing total time by a factor of k (minus overhead).

### Assumptions for Parallel Execution
- **Thread Limit**: 10 concurrent jobs (adjustable in script).
- **PowerShell Version**: 5.1+ with job support.
- **Overhead**: File writing/reading (~seconds), job startup/teardown (~1-2 seconds total).
- **Network Sharing**: Parallel requests may share bandwidth; very high parallelism could hit rate limits.

## Estimated Execution Times

### Single-Threaded Estimates (Original)
| Number of URLs | Estimated Time | Notes |
|----------------|----------------|-------|
| 1,000 | 30-35 minutes | Sequential processing. |
| 5,000 | 2.5-3 hours | Feasible but slow. |
| 10,000 | 5.5-6 hours | Requires patience. |
| 50,000 | 27-28 hours | Multi-day run. |
| 100,000 | 55-56 hours | Extremely long. |

### Parallel Estimates (Updated, k=50 threads, optimized)
| Number of URLs | Estimated Time | Speedup Factor | Notes |
|----------------|----------------|----------------|-------|
| 1,000 | 1-2 minutes | ~15-30x | Very fast for small batches. |
| 5,000 | 5-10 minutes | ~15-30x | Efficient batch processing. |
| 10,000 | 10-20 minutes | ~15-30x | Under 30 minutes. |
| 50,000 | 50-100 minutes | ~15-30x | 1-2 hours; practical. |
| 100,000 | 1.5-3 hours | ~15-30x | Manageable; test network limits. |

### Factors Affecting Performance
- **Parallelism Level (k)**: Higher k (e.g., 20) can further reduce time but may strain network/system.
- **Network Speed**: Faster connections amplify parallel benefits.
- **Failure Rate**: Retries add time; parallel handles failures better.
- **System Resources**: CPU/RAM for jobs; Excel waits for completion.
- **Rate Limiting**: Some servers may throttle parallel requests.

### Recommendations
- **Tuning**: Adjust `MaxThreads` in `URLChecker.ps1` based on network capacity (e.g., 20-100).
- **Large Datasets**: For 100k+ URLs, consider splitting runs or using dedicated servers.
- **Fallback**: Original `CheckURLStatus` function retained for single-threaded mode if needed.
- **Testing**: Start with small batches to verify parallel behavior.

### Further Optimizations for 100,000+ URLs
- **PowerShell 7**: Upgrade to PS 7 for `ForEach-Object -Parallel` (even faster than runspaces).
- **Hardware**: Use multi-core CPU, high-speed internet (fiber preferred).
- **Rate Limiting**: Monitor for server throttling; add delays if needed.
- **Async Processing**: For extreme scale, consider .NET HttpClient in a compiled script.
- **Batch Splitting**: Divide into multiple PS runs if memory/network limits hit.
- **Cloud Resources**: Run on Azure VMs with better bandwidth for massive datasets.

The optimized parallel version makes 100,000+ URLs feasible, reducing time from days to hours.