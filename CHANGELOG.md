# Changelog

All notable changes to this project are documented here. Format based on [Keep a Changelog](https://keepachangelog.com/).

## [Unreleased]

### Fixed
- Snapshot age (vSnapshot *Age (days)* and the vHealth snapshot check) is now measured from the inventory's collection time, as documented, instead of the time the workbook is built. A saved inventory processed with `--from-json` gives the same ages and findings on any date, and the bundled sample gives `vHealth=50` again.

## [1.0.0] - 2026-10-20

First public release. Verified against a live HPE VM Essentials 9.1.0.1 manager.

### Added
- Username/password (OAuth) and API token authentication; `--insecure` for self-signed certificates.
- Paged collection of clouds, groups, clusters, servers (hosts and VMs), instances, networks, datastores and snapshots.
- Tabs: vInfo, vCPU, vMemory, vDisk, vNetwork, vSnapshot, vHost, vCluster, vDatastore, vNetworks, vInstance, vHealth, vSource.
- vDatastore includes cluster Storage-tab datastores (GFS2 shared LUN, NFS, local) merged with cloud datastores, with provisioned/overcommit %, host and cluster columns, NFS server (*Address*), NFS export or LUN device (*Source*) and mount path.
- vHost hardware columns from the 9.1 API: *CPU Model*, *CPU MHz*, *# Sockets*, *Vendor*, *Model*, and a *Maintenance* column.
- Health checks for host state, host maintenance mode, host memory, datastore free space, snapshot age, VM state and missing agents.
- Usage stats: servers whose `/api/servers` list record has no `stats.usedMemory` are read individually from `GET /api/servers/{id}`, with a fallback to the instance's stats for single-VM instances. `--no-stats` skips this.
- `--raw` sheets (raw_servers, raw_instances, raw_datastores), `--dump-json` and `--from-json`.
- Environment variable configuration (MORPHEUS_URL, MORPHEUS_USER, MORPHEUS_PASS, MORPHEUS_TOKEN).
- `--version`.
- Credential redaction: `apiKey`, `sshPassword`, `sshPasswordHash`, service passwords/tokens and similar fields are replaced with `<redacted>` in raw sheets, `--dump-json` output and `--from-json` input.
- Clear error messages and a non-zero exit for TLS failures, unreachable managers, bad credentials, a locked output file and invalid JSON input. A connection failure now stops the run instead of producing an empty workbook.
- `examples/generate_sample.py`, which generates the fictional sample inventory (4 clusters, 17 hosts, 109 VMs) in 9.1 record shapes.
- Documentation: workbook recipes (`docs/USING-THE-WORKBOOK.md`), Excel screenshots of the main tabs, compatibility notes, and documented redaction and error handling.

### Fixed (found when testing against VME 9.1.0.1)
- Hosts are detected by `computeServerType.format = vmHypervisor`. 9.1 has no `vmHypervisor` flag, so a host with no VMs used to be listed as a VM.
- Host cluster comes from `server.serverGroup` when present.
- vHost *CPU Model*, *Vendor* and *Model* were always blank; they now read `hardwareCpuModel`, `hardwareProductVendor` and `hardwareProductName`, with trailing newlines trimmed. *# Sockets* reads `maxSockets`. The unpopulated *Serial* column has been removed.
- vDatastore *Type* showed `generic` for GFS2 and NFS datastores; it now uses `datastoreType.name`.
- vHost *OS* shows the host platform (`ubuntu`) rather than the generic `linux`.
- vInfo *Managed* uses `computeServerType.managed`, so discovered VMs such as the manager appliance show `False`.
- vInstance *IP Addresses* no longer shows the `0.0.0.0` placeholder.
- Powered-off VMs (status `stopped`) no longer get a duplicate *Status: stopped* warning in vHealth.
- *CPU Usage %* is rounded to one decimal place in vCPU and vHost.
- Column widths leave room for the AutoFilter button, so headers are no longer cut off.
- vDisk *Unit #* and vSnapshot *Instance ID* are written as numbers rather than text.
- vHealth sorts hosts, datastores and snapshots ahead of VMs within each severity, so infrastructure findings aren't buried under per-VM warnings.
- vSnapshot *State* was blank; 9.1 reports it in `status`. Added *Size MiB* and *Memory* (memory snapshot) columns.
