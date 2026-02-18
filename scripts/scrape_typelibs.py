#!/usr/bin/env python
"""Scrape ALL Solid Edge type libraries into structured JSON + summary markdown.

Discovers every .tlb file under the Solid Edge install directory, loads each
with pythoncom.LoadTypeLib, and extracts every enum, dispatch/dual interface,
coclass, record, and alias into a single reference JSON file.

Usage:
    uv run python scripts/scrape_typelibs.py
    uv run python scripts/scrape_typelibs.py --install-dir "C:\\Custom\\Path"
    uv run python scripts/scrape_typelibs.py --output reference/typelib_dump.json

Outputs:
    reference/typelib_dump.json   - Full structured dump
    reference/typelib_summary.md  - Human-readable summary
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import UTC, datetime
from pathlib import Path

try:
    import pythoncom
except ImportError:
    print("ERROR: pythoncom not available. Install pywin32: pip install pywin32")
    sys.exit(1)


# ---------------------------------------------------------------------------
# VT type code → human-readable name
# ---------------------------------------------------------------------------

VT_NAMES: dict[int, str] = {
    pythoncom.VT_EMPTY: "VT_EMPTY",
    pythoncom.VT_NULL: "VT_NULL",
    pythoncom.VT_I2: "VT_I2",
    pythoncom.VT_I4: "VT_I4",
    pythoncom.VT_R4: "VT_R4",
    pythoncom.VT_R8: "VT_R8",
    pythoncom.VT_CY: "VT_CY",
    pythoncom.VT_DATE: "VT_DATE",
    pythoncom.VT_BSTR: "VT_BSTR",
    pythoncom.VT_DISPATCH: "VT_DISPATCH",
    pythoncom.VT_ERROR: "VT_ERROR",
    pythoncom.VT_BOOL: "VT_BOOL",
    pythoncom.VT_VARIANT: "VT_VARIANT",
    pythoncom.VT_UNKNOWN: "VT_UNKNOWN",
    pythoncom.VT_DECIMAL: "VT_DECIMAL",
    pythoncom.VT_I1: "VT_I1",
    pythoncom.VT_UI1: "VT_UI1",
    pythoncom.VT_UI2: "VT_UI2",
    pythoncom.VT_UI4: "VT_UI4",
    pythoncom.VT_I8: "VT_I8",
    pythoncom.VT_UI8: "VT_UI8",
    pythoncom.VT_INT: "VT_INT",
    pythoncom.VT_UINT: "VT_UINT",
    pythoncom.VT_VOID: "VT_VOID",
    pythoncom.VT_HRESULT: "VT_HRESULT",
    pythoncom.VT_PTR: "VT_PTR",
    pythoncom.VT_SAFEARRAY: "VT_SAFEARRAY",
    pythoncom.VT_CARRAY: "VT_CARRAY",
    pythoncom.VT_USERDEFINED: "VT_USERDEFINED",
    pythoncom.VT_LPSTR: "VT_LPSTR",
    pythoncom.VT_LPWSTR: "VT_LPWSTR",
}

# TKIND constants
TKIND_ENUM = 0
TKIND_RECORD = 1
TKIND_MODULE = 2
TKIND_INTERFACE = 3
TKIND_DISPATCH = 4
TKIND_COCLASS = 5
TKIND_ALIAS = 6
TKIND_UNION = 7

TKIND_NAMES = {
    TKIND_ENUM: "enum",
    TKIND_RECORD: "record",
    TKIND_MODULE: "module",
    TKIND_INTERFACE: "interface",
    TKIND_DISPATCH: "dispatch",
    TKIND_COCLASS: "coclass",
    TKIND_ALIAS: "alias",
    TKIND_UNION: "union",
}

# INVOKEKIND constants
INVOKE_FUNC = 1
INVOKE_PROPERTYGET = 2
INVOKE_PROPERTYPUT = 4
INVOKE_PROPERTYPUTREF = 8

# FUNCFLAGS
FUNCFLAG_FRESTRICTED = 1
FUNCFLAG_FSOURCE = 2
FUNCFLAG_FHIDDEN = 64

# PARAMFLAGS
PARAMFLAG_FIN = 1
PARAMFLAG_FOUT = 2
PARAMFLAG_FRETVAL = 8
PARAMFLAG_FOPT = 16
PARAMFLAG_FHASDEFAULT = 32


# ---------------------------------------------------------------------------
# Type resolution
# ---------------------------------------------------------------------------


def resolve_type_desc(type_desc, typeinfo) -> str:
    """Resolve a TYPEDESC to a human-readable type string.

    pythoncom represents TYPEDESCs as:
    - int: simple VT type (e.g., 3 = VT_I4)
    - (vt, inner): compound type where vt is VT_PTR/VT_SAFEARRAY/VT_USERDEFINED
      and inner is a nested type_desc (for PTR/SAFEARRAY) or href (for USERDEFINED)
    """
    if type_desc is None:
        return "void"

    # Simple VT type: just an integer
    if isinstance(type_desc, int):
        return VT_NAMES.get(type_desc, f"VT_{type_desc}")

    # Compound type: (vt, inner)
    if not isinstance(type_desc, tuple) or len(type_desc) != 2:
        return f"unknown({type_desc!r})"

    vt, inner = type_desc

    if vt == pythoncom.VT_PTR:
        inner_str = resolve_type_desc(inner, typeinfo)
        return f"{inner_str}*"

    if vt == pythoncom.VT_SAFEARRAY:
        inner_str = resolve_type_desc(inner, typeinfo)
        return f"SAFEARRAY({inner_str})"

    if vt == pythoncom.VT_CARRAY:
        inner_str = resolve_type_desc(inner, typeinfo)
        return f"CARRAY({inner_str})"

    if vt == pythoncom.VT_USERDEFINED:
        href = inner
        try:
            ref_info = typeinfo.GetRefTypeInfo(href)
            ref_name = ref_info.GetDocumentation(-1)[0]
            return ref_name
        except Exception:
            return f"UserDefined(href={href})"

    return f"VT_{vt}({inner!r})"


def param_flags_str(flags: int) -> str:
    """Convert parameter flags bitmask to string."""
    parts = []
    if flags & PARAMFLAG_FIN:
        parts.append("in")
    if flags & PARAMFLAG_FOUT:
        parts.append("out")
    if flags & PARAMFLAG_FRETVAL:
        parts.append("retval")
    if flags & PARAMFLAG_FOPT:
        parts.append("optional")
    if flags & PARAMFLAG_FHASDEFAULT:
        parts.append("hasdefault")
    return ",".join(parts) if parts else "in"


# ---------------------------------------------------------------------------
# Extraction functions
# ---------------------------------------------------------------------------


def extract_enum(typeinfo, typeattr) -> dict[str, int]:
    """Extract all members of an enum."""
    members = {}
    for i in range(typeattr.cVars):
        try:
            vd = typeinfo.GetVarDesc(i)
            names = typeinfo.GetNames(vd.memid)
            name = names[0] if names else f"var_{i}"
            value = vd.value
            # Convert to int if possible (some are already int)
            if isinstance(value, (int, bool)):
                members[name] = int(value)
            else:
                members[name] = value
        except Exception:
            continue
    return members


def extract_record(typeinfo, typeattr) -> dict:
    """Extract fields from a record (struct)."""
    fields = {}
    for i in range(typeattr.cVars):
        try:
            vd = typeinfo.GetVarDesc(i)
            names = typeinfo.GetNames(vd.memid)
            name = names[0] if names else f"field_{i}"
            # elemdescVar is (type_desc, flags, default)
            type_str = resolve_type_desc(vd.elemdescVar[0], typeinfo)
            fields[name] = type_str
        except Exception:
            continue
    return {"fields": fields} if fields else {}


def extract_interface(typeinfo, typeattr) -> dict:
    """Extract methods and properties from a dispatch/dual interface."""
    methods = {}
    properties = {}
    iid = str(typeattr.iid)

    for i in range(typeattr.cFuncs):
        try:
            fd = typeinfo.GetFuncDesc(i)
        except Exception:
            continue

        # Skip restricted/hidden members
        if fd.wFuncFlags & FUNCFLAG_FRESTRICTED:
            continue

        try:
            names = typeinfo.GetNames(fd.memid)
        except Exception:
            continue

        if not names:
            continue

        func_name = names[0]
        invkind = fd.invkind

        # Resolve return type - fd.rettype is (type_desc, flags, default)
        ret_type = resolve_type_desc(fd.rettype[0], typeinfo)

        # Build params list
        params = []
        param_names = names[1:]  # first name is the function name
        for pi in range(len(fd.args)):
            arg_desc = fd.args[pi]
            # arg_desc is (type_desc, param_flags, [default_value])
            type_desc = arg_desc[0]
            pflags = arg_desc[1] if len(arg_desc) > 1 else PARAMFLAG_FIN

            pname = param_names[pi] if pi < len(param_names) else f"p{pi}"
            ptype = resolve_type_desc(type_desc, typeinfo)
            pflags_s = param_flags_str(pflags)

            param_info = {
                "name": pname,
                "type": ptype,
                "flags": pflags_s,
            }

            # Include default value if present
            if len(arg_desc) > 2 and arg_desc[2] is not None:
                default_val = arg_desc[2]
                # Make JSON-serializable
                if isinstance(default_val, (int, float, str, bool)):
                    param_info["default"] = default_val
                else:
                    param_info["default"] = str(default_val)

            params.append(param_info)

        if invkind == INVOKE_FUNC:
            methods[func_name] = {
                "dispid": fd.memid,
                "returns": ret_type,
                "params": params,
            }
        elif invkind == INVOKE_PROPERTYGET:
            prop_entry = properties.setdefault(func_name, {})
            prop_entry["type"] = ret_type
            prop_entry["dispid"] = fd.memid
            access = prop_entry.get("access", "")
            prop_entry["access"] = "get/put" if "put" in access else "get"
            if params:
                prop_entry["index_params"] = params
        elif invkind in (INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF):
            prop_entry = properties.setdefault(func_name, {})
            prop_entry["dispid"] = fd.memid
            access = prop_entry.get("access", "")
            if invkind == INVOKE_PROPERTYPUTREF:
                prop_entry["access"] = "get/putref" if "get" in access else "putref"
            else:
                prop_entry["access"] = "get/put" if "get" in access else "put"
            if params and "type" not in prop_entry:
                prop_entry["type"] = params[-1]["type"]

    result = {"iid": iid}

    # Inherited interfaces
    for ii in range(typeattr.cImplTypes):
        try:
            href = typeinfo.GetRefTypeOfImplType(ii)
            ref_info = typeinfo.GetRefTypeInfo(href)
            ref_name = ref_info.GetDocumentation(-1)[0]
            result.setdefault("inherits", []).append(ref_name)
        except Exception:
            pass

    if methods:
        result["methods"] = methods
    if properties:
        result["properties"] = properties

    return result


def extract_coclass(typeinfo, typeattr) -> dict:
    """Extract implemented interfaces from a coclass."""
    clsid = str(typeattr.iid)
    interfaces = []

    for i in range(typeattr.cImplTypes):
        try:
            flags = typeinfo.GetImplTypeFlags(i)
            href = typeinfo.GetRefTypeOfImplType(i)
            ref_info = typeinfo.GetRefTypeInfo(href)
            ref_name = ref_info.GetDocumentation(-1)[0]

            iface_info = {"name": ref_name}
            flag_parts = []
            if flags & 1:  # IMPLTYPEFLAG_FDEFAULT
                flag_parts.append("default")
            if flags & 2:  # IMPLTYPEFLAG_FSOURCE
                flag_parts.append("source")
            if flag_parts:
                iface_info["flags"] = ",".join(flag_parts)
            interfaces.append(iface_info)
        except Exception:
            continue

    return {"clsid": clsid, "interfaces": interfaces}


def extract_module(typeinfo, typeattr) -> dict:
    """Extract constants and functions from a module."""
    constants = {}
    functions = {}

    # Module variables (constants)
    for i in range(typeattr.cVars):
        try:
            vd = typeinfo.GetVarDesc(i)
            names = typeinfo.GetNames(vd.memid)
            name = names[0] if names else f"const_{i}"
            value = vd.value
            if isinstance(value, (int, float, str, bool)):
                constants[name] = value
            else:
                constants[name] = str(value)
        except Exception:
            continue

    # Module functions
    for i in range(typeattr.cFuncs):
        try:
            fd = typeinfo.GetFuncDesc(i)
            names = typeinfo.GetNames(fd.memid)
            if not names:
                continue
            func_name = names[0]
            ret_type = resolve_type_desc(fd.rettype[0], typeinfo)
            params = []
            param_names = names[1:]
            for pi in range(len(fd.args)):
                arg_desc = fd.args[pi]
                type_desc = arg_desc[0]
                pflags = arg_desc[1] if len(arg_desc) > 1 else PARAMFLAG_FIN
                pname = param_names[pi] if pi < len(param_names) else f"p{pi}"
                ptype = resolve_type_desc(type_desc, typeinfo)
                params.append(
                    {
                        "name": pname,
                        "type": ptype,
                        "flags": param_flags_str(pflags),
                    }
                )
            functions[func_name] = {
                "returns": ret_type,
                "params": params,
            }
        except Exception:
            continue

    result = {}
    if constants:
        result["constants"] = constants
    if functions:
        result["functions"] = functions
    return result


# ---------------------------------------------------------------------------
# Type library scraping
# ---------------------------------------------------------------------------


def find_typelibs(install_dir: str) -> list[Path]:
    """Find all .tlb files under the Solid Edge install directory."""
    root = Path(install_dir)
    if not root.exists():
        print(f"WARNING: Install directory does not exist: {install_dir}")
        return []
    tlbs = sorted(root.rglob("*.tlb"))
    return tlbs


def scrape_typelib(tlb_path: Path) -> dict:
    """Load and extract all type information from a single .tlb file."""
    tlib = pythoncom.LoadTypeLib(str(tlb_path))
    lib_attr = tlib.GetLibAttr()
    lib_doc = tlib.GetDocumentation(-1)

    result = {
        "file": tlb_path.name,
        "path": str(tlb_path),
        "guid": str(lib_attr[0]),
        "description": lib_doc[1] or "",
        "version": f"{lib_attr[3]}.{lib_attr[4]}",
        "lcid": lib_attr[1],
    }

    enums = {}
    interfaces = {}
    coclasses = {}
    records = {}
    modules = {}
    aliases = {}

    count = tlib.GetTypeInfoCount()
    for i in range(count):
        try:
            tkind = tlib.GetTypeInfoType(i)
            ti = tlib.GetTypeInfo(i)
            doc = tlib.GetDocumentation(i)
            name = doc[0]
            ta = ti.GetTypeAttr()
        except Exception:
            continue

        try:
            if tkind == TKIND_ENUM:
                members = extract_enum(ti, ta)
                if members:
                    enums[name] = members

            elif tkind in (TKIND_DISPATCH, TKIND_INTERFACE):
                iface = extract_interface(ti, ta)
                iface["kind"] = TKIND_NAMES.get(tkind, str(tkind))
                interfaces[name] = iface

            elif tkind == TKIND_COCLASS:
                coclasses[name] = extract_coclass(ti, ta)

            elif tkind == TKIND_RECORD:
                rec = extract_record(ti, ta)
                if rec:
                    records[name] = rec

            elif tkind == TKIND_MODULE:
                mod = extract_module(ti, ta)
                if mod:
                    modules[name] = mod

            elif tkind == TKIND_ALIAS:
                target = resolve_type_desc(ta.tdescAlias, ti)
                aliases[name] = target

        except Exception as e:
            # Log but continue
            result.setdefault("_errors", []).append(
                f"Error extracting {TKIND_NAMES.get(tkind, '?')} '{name}': {e}"
            )

    if enums:
        result["enums"] = enums
    if interfaces:
        result["interfaces"] = interfaces
    if coclasses:
        result["coclasses"] = coclasses
    if records:
        result["records"] = records
    if modules:
        result["modules"] = modules
    if aliases:
        result["aliases"] = aliases

    return result


# ---------------------------------------------------------------------------
# Summary generation
# ---------------------------------------------------------------------------


def generate_summary(data: dict, output_path: Path) -> None:
    """Generate a human-readable markdown summary."""
    meta = data["metadata"]
    typelibs = data["typelibs"]

    lines = []
    lines.append("# Solid Edge Type Library Reference")
    lines.append("")
    lines.append(
        f"Generated: {meta['scraped_at'][:19]} | "
        f"SE {meta['solid_edge_version']} | "
        f"{meta['typelib_count']} type libraries"
    )
    lines.append("")

    # Global statistics
    total_enums = 0
    total_enum_values = 0
    total_interfaces = 0
    total_methods = 0
    total_properties = 0
    total_coclasses = 0

    for _tlb_name, tlb_data in typelibs.items():
        enums = tlb_data.get("enums", {})
        ifaces = tlb_data.get("interfaces", {})
        coclasses = tlb_data.get("coclasses", {})
        total_enums += len(enums)
        total_enum_values += sum(len(v) for v in enums.values())
        total_interfaces += len(ifaces)
        for iface in ifaces.values():
            total_methods += len(iface.get("methods", {}))
            total_properties += len(iface.get("properties", {}))
        total_coclasses += len(coclasses)

    lines.append("## Global Statistics")
    lines.append("")
    lines.append("| Metric | Count |")
    lines.append("|--------|-------|")
    lines.append(f"| Type libraries | {meta['typelib_count']} |")
    lines.append(f"| Enums | {total_enums} |")
    lines.append(f"| Enum values | {total_enum_values} |")
    lines.append(f"| Interfaces (dispatch+vtable) | {total_interfaces} |")
    lines.append(f"| Methods | {total_methods} |")
    lines.append(f"| Properties | {total_properties} |")
    lines.append(f"| CoClasses | {total_coclasses} |")
    lines.append("")

    # Overview table
    lines.append("## Type Library Overview")
    lines.append("")
    lines.append("| File | Description | Enums | Interfaces | CoClasses |")
    lines.append("|------|-------------|-------|------------|-----------|")
    for tlb_name, tlb_data in sorted(typelibs.items()):
        desc = tlb_data.get("description", "")[:60]
        ne = len(tlb_data.get("enums", {}))
        ni = len(tlb_data.get("interfaces", {}))
        nc = len(tlb_data.get("coclasses", {}))
        lines.append(f"| {tlb_name} | {desc} | {ne} | {ni} | {nc} |")
    lines.append("")

    # Per-library details
    for tlb_name, tlb_data in sorted(typelibs.items()):
        desc = tlb_data.get("description", tlb_name)
        lines.append("---")
        lines.append(f"## {tlb_name}")
        guid = tlb_data.get("guid", "?")
        ver = tlb_data.get("version", "?")
        lines.append(f"**{desc}** (GUID: `{guid}`, v{ver})")
        lines.append("")

        # Enums
        enums = tlb_data.get("enums", {})
        if enums:
            lines.append(f"### Enums ({len(enums)})")
            lines.append("")
            for enum_name in sorted(enums.keys()):
                members = enums[enum_name]
                # Show first few members inline
                member_strs = [
                    f"{k}={v}"
                    for k, v in sorted(
                        members.items(),
                        key=lambda x: x[1] if isinstance(x[1], (int, float)) else 0,
                    )
                ]
                preview = ", ".join(member_strs[:5])
                if len(member_strs) > 5:
                    preview += f", ... ({len(member_strs)} total)"
                lines.append(f"- **{enum_name}** [{len(members)}]: {preview}")
            lines.append("")

        # Interfaces
        ifaces = tlb_data.get("interfaces", {})
        if ifaces:
            lines.append(f"### Interfaces ({len(ifaces)})")
            lines.append("")
            lines.append("| Interface | Kind | Methods | Properties |")
            lines.append("|-----------|------|---------|------------|")
            for iface_name in sorted(ifaces.keys()):
                iface = ifaces[iface_name]
                kind = iface.get("kind", "?")
                nm = len(iface.get("methods", {}))
                np = len(iface.get("properties", {}))
                lines.append(f"| {iface_name} | {kind} | {nm} | {np} |")
            lines.append("")

        # CoClasses
        coclasses = tlb_data.get("coclasses", {})
        if coclasses:
            lines.append(f"### CoClasses ({len(coclasses)})")
            lines.append("")
            for cc_name in sorted(coclasses.keys()):
                cc = coclasses[cc_name]
                iface_names = [i["name"] for i in cc.get("interfaces", [])]
                clsid = cc.get("clsid", "?")
                ifaces_str = ", ".join(iface_names)
                lines.append(f"- **{cc_name}** (`{clsid}`): {ifaces_str}")
            lines.append("")

        # Aliases
        aliases = tlb_data.get("aliases", {})
        if aliases:
            lines.append(f"### Aliases ({len(aliases)})")
            lines.append("")
            for alias_name in sorted(aliases.keys()):
                lines.append(f"- {alias_name} = {aliases[alias_name]}")
            lines.append("")

    output_path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

DEFAULT_INSTALL_DIR = r"C:\Program Files\Siemens\Solid Edge 2026"


def main():
    parser = argparse.ArgumentParser(
        description="Scrape Solid Edge type libraries into JSON + markdown reference"
    )
    parser.add_argument(
        "--install-dir",
        default=DEFAULT_INSTALL_DIR,
        help=f"Solid Edge install directory (default: {DEFAULT_INSTALL_DIR})",
    )
    parser.add_argument(
        "--output",
        default="reference/typelib_dump.json",
        help="Output JSON file path (default: reference/typelib_dump.json)",
    )
    parser.add_argument(
        "--summary",
        default="reference/typelib_summary.md",
        help="Output summary markdown path (default: reference/typelib_summary.md)",
    )
    args = parser.parse_args()

    output_path = Path(args.output)
    summary_path = Path(args.summary)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.parent.mkdir(parents=True, exist_ok=True)

    print("Solid Edge Type Library Scraper")
    print(f"{'=' * 50}")
    print(f"Install dir: {args.install_dir}")
    print(f"Output JSON: {output_path}")
    print(f"Output MD:   {summary_path}")
    print()

    # Discover type libraries
    print("Discovering .tlb files...")
    tlb_files = find_typelibs(args.install_dir)
    if not tlb_files:
        print("ERROR: No .tlb files found. Is Solid Edge installed?")
        sys.exit(1)

    print(f"Found {len(tlb_files)} type libraries")
    print()

    # Scrape each
    typelibs = {}
    failures = []
    t0 = time.time()

    for i, tlb_path in enumerate(tlb_files, 1):
        rel = tlb_path.relative_to(args.install_dir)
        print(f"  [{i:2d}/{len(tlb_files)}] {rel} ... ", end="", flush=True)
        try:
            result = scrape_typelib(tlb_path)
            # Use relative path as key for readability
            key = str(rel).replace("\\", "/")
            typelibs[key] = result

            # Summary counts
            ne = len(result.get("enums", {}))
            ni = len(result.get("interfaces", {}))
            nc = len(result.get("coclasses", {}))
            errors = result.get("_errors", [])
            status = f"{ne} enums, {ni} interfaces, {nc} coclasses"
            if errors:
                status += f" ({len(errors)} errors)"
            print(status)
        except Exception as e:
            failures.append((str(rel), str(e)))
            print(f"FAILED: {e}")

    elapsed = time.time() - t0
    print()
    print(f"Scraped {len(typelibs)}/{len(tlb_files)} type libraries in {elapsed:.1f}s")

    if failures:
        print(f"\nFailed ({len(failures)}):")
        for name, err in failures:
            print(f"  {name}: {err}")

    # Build output
    data = {
        "metadata": {
            "scraped_at": datetime.now(UTC).isoformat(),
            "solid_edge_version": "2026",
            "install_path": args.install_dir,
            "typelib_count": len(typelibs),
            "failures": len(failures),
        },
        "typelibs": typelibs,
    }

    # Write JSON
    print(f"\nWriting JSON to {output_path} ...")
    output_path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False, default=str),
        encoding="utf-8",
    )
    json_size = output_path.stat().st_size
    print(f"  {json_size:,} bytes")

    # Write summary
    print(f"Writing summary to {summary_path} ...")
    generate_summary(data, summary_path)
    md_size = summary_path.stat().st_size
    print(f"  {md_size:,} bytes")

    # Print global stats
    print()
    print("=" * 50)
    print("SUMMARY")
    print("=" * 50)
    total_enums = sum(len(t.get("enums", {})) for t in typelibs.values())
    total_enum_values = sum(
        sum(len(v) for v in t.get("enums", {}).values()) for t in typelibs.values()
    )
    total_interfaces = sum(len(t.get("interfaces", {})) for t in typelibs.values())
    total_methods = sum(
        sum(len(iface.get("methods", {})) for iface in t.get("interfaces", {}).values())
        for t in typelibs.values()
    )
    total_properties = sum(
        sum(len(iface.get("properties", {})) for iface in t.get("interfaces", {}).values())
        for t in typelibs.values()
    )
    total_coclasses = sum(len(t.get("coclasses", {})) for t in typelibs.values())
    total_aliases = sum(len(t.get("aliases", {})) for t in typelibs.values())

    print(f"  Type libraries: {len(typelibs)}")
    print(f"  Enums:          {total_enums} ({total_enum_values} values)")
    print(f"  Interfaces:     {total_interfaces}")
    print(f"  Methods:        {total_methods}")
    print(f"  Properties:     {total_properties}")
    print(f"  CoClasses:      {total_coclasses}")
    print(f"  Aliases:        {total_aliases}")
    print()


if __name__ == "__main__":
    main()
