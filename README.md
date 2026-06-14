# Sentinel VB6

Sentinel is a VB6-based Cafe Management System developed during the early Windows LAN and internet café era for operational PC rental and workstation management environments.

Internet cafés proliferated during the 1990s and early 2000s because mobile internet and affordable home DSL connections were still uncommon. Online multiplayer gaming also existed primarily through LAN-based gaming environments. Due to the growing demand for computer rental services, centralized timer systems became necessary to monitor and track billable client usage time.

At the time of development, Windows 95 and later Windows 98 systems were dominant. However, the networking stack of that era was relatively unstable and limited for environments consisting of one centralized server coordinating more than a hundred client workstations simultaneously.

The system was designed around practical business and operational requirements including session monitoring, time recording, workstation access control, reporting, security restrictions, and centralized client management across local area networks.

Originally deployed using a client/server architecture, Sentinel included operational features such as remote workstation control, automated timeout handling, audit logging, encrypted data storage, desktop locking, reporting, and multi-level user security. The system was successfully used in real-world computer rental shop environments and was designed with emphasis on simplicity, maintainability, operational reliability, and low network overhead.

The project reflects practical software engineering approaches from the Windows 98 / Windows ME networking era prior to modern cloud and web-based operational platforms. Much of the architecture evolved incrementally through real-world deployment, maintenance, troubleshooting, and operational use.

This repository is preserved as a historical engineering artifact representing an earlier generation of operational business systems development and workstation management software.

## Server Architecture

The Server application is divided into logical operational components:

```text
Server
 ├── Session Management
 ├── Billing
 ├── Reporting
 ├── Security
 └── Client Coordination
```

## Client Architecture

The Client application responds to commands issued by the server and is designed to remain partially autonomous during disconnected or degraded network conditions.

```text
Client
 ├── Session Monitoring
 ├── Workstation Locking
 ├── Timeout Handling
 ├── Command Processing
 └── Local Failover Logic
```
