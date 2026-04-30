from __future__ import annotations

import importlib.util
import json
import logging
import shutil
from pathlib import Path
from urllib.parse import urlparse

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import pandas as pd
import seaborn as sns

from .models import ErrorRecord, LinkRecord, PageRecord, PdfRecord
from .parser import section_from_url

LOGGER = logging.getLogger("moj_isa_crawler")


def _frame(rows: list[dict[str, object]]) -> pd.DataFrame:
    return pd.DataFrame(rows)


def build_stats_frames(
    *,
    pages: list[PageRecord],
    pdfs: list[PdfRecord],
    links: list[LinkRecord],
    errors: list[ErrorRecord],
) -> dict[str, pd.DataFrame]:
    page_df = pd.DataFrame([p.as_dict() for p in pages])
    pdf_df = pd.DataFrame([p.as_dict() for p in pdfs])
    link_df = pd.DataFrame([l.as_dict() for l in links])
    error_df = pd.DataFrame([e.as_dict() for e in errors])

    unique_pdf_urls = set(pdf_df["pdf_url"]) if not pdf_df.empty and "pdf_url" in pdf_df else set()
    downloaded_unique = set(pdf_df.loc[pdf_df.get("downloaded", False) == True, "pdf_url"]) if not pdf_df.empty and "downloaded" in pdf_df else set()  # noqa: E712
    internal_links = link_df[link_df["category"] == "internal_page"] if not link_df.empty and "category" in link_df else pd.DataFrame()
    external_links = link_df[link_df["category"] == "external"] if not link_df.empty and "category" in link_df else pd.DataFrame()

    stats_rows = [
        {"metric": "pages", "value": len(pages)},
        {"metric": "sections", "value": int(page_df["section"].nunique()) if not page_df.empty and "section" in page_df else 0},
        {"metric": "links", "value": len(links)},
        {"metric": "internal_page_links", "value": len(internal_links)},
        {"metric": "external_links", "value": len(external_links)},
        {"metric": "pdf_references", "value": len(pdfs)},
        {"metric": "unique_pdf_urls", "value": len(unique_pdf_urls)},
        {"metric": "downloaded_unique_pdf_urls", "value": len(downloaded_unique)},
        {"metric": "errors", "value": len(errors)},
        {"metric": "max_depth", "value": int(page_df["depth"].max()) if not page_df.empty and "depth" in page_df else 0},
        {"metric": "mean_text_length", "value": float(np.round(page_df["text_length"].mean(), 2)) if not page_df.empty and "text_length" in page_df else 0},
    ]

    if not page_df.empty:
        depth_stats = (
            page_df.groupby("depth", dropna=False)
            .agg(pages=("url", "count"), mean_text_length=("text_length", "mean"), pdf_links=("pdf_link_count", "sum"), internal_page_links=("internal_page_link_count", "sum"))
            .reset_index()
        )
        depth_stats["mean_text_length"] = depth_stats["mean_text_length"].round(2)
        section_stats = (
            page_df.groupby("section", dropna=False)
            .agg(pages=("url", "count"), pdf_links=("pdf_link_count", "sum"), internal_page_links=("internal_page_link_count", "sum"), mean_text_length=("text_length", "mean"))
            .reset_index()
        )
        section_stats["mean_text_length"] = section_stats["mean_text_length"].round(2)
        section_stats = section_stats.sort_values(["pages", "pdf_links"], ascending=False)
        top_pages = page_df.sort_values(["pdf_link_count", "link_count", "text_length"], ascending=False)[
            ["order", "depth", "section", "title", "url", "pdf_link_count", "link_count", "text_length"]
        ].head(50)
    else:
        depth_stats = pd.DataFrame(columns=["depth", "pages", "mean_text_length", "pdf_links", "internal_page_links"])
        section_stats = pd.DataFrame(columns=["section", "pages", "pdf_links", "internal_page_links", "mean_text_length"])
        top_pages = pd.DataFrame(columns=["order", "depth", "section", "title", "url", "pdf_link_count", "link_count", "text_length"])

    if not pdf_df.empty:
        pdf_stats = (
            pdf_df.groupby("source_section", dropna=False)
            .agg(pdf_references=("pdf_url", "count"), unique_pdf_urls=("pdf_url", "nunique"), downloaded=("downloaded", "sum"), errored=("error", lambda values: int(sum(bool(v) for v in values))))
            .reset_index()
            .sort_values(["pdf_references", "unique_pdf_urls"], ascending=False)
        )
    else:
        pdf_stats = pd.DataFrame(columns=["source_section", "pdf_references", "unique_pdf_urls", "downloaded", "errored"])

    if not error_df.empty:
        error_stats = (
            error_df.groupby(["phase", "exception_type"], dropna=False)
            .agg(errors=("url", "count"))
            .reset_index()
            .sort_values("errors", ascending=False)
        )
    else:
        error_stats = pd.DataFrame(columns=["phase", "exception_type", "errors"])

    return {
        "Stats": _frame(stats_rows),
        "DepthStats": depth_stats,
        "SectionStats": section_stats,
        "PdfStats": pdf_stats,
        "ErrorStats": error_stats,
        "TopPages": top_pages,
    }


def _save_bar(frame: pd.DataFrame, *, x: str, y: str, title: str, output: Path, xlabel: str = "", ylabel: str = "") -> None:
    plt.figure(figsize=(12, 7))
    sns.barplot(data=frame, x=x, y=y, color="#4C78A8")
    plt.title(title)
    plt.xlabel(xlabel or x)
    plt.ylabel(ylabel or y)
    plt.xticks(rotation=35, ha="right")
    plt.tight_layout()
    plt.savefig(output, dpi=160)
    plt.close()


def _safe_graph_name(url: str) -> str:
    parsed = urlparse(url)
    name = parsed.path.rstrip("/").rsplit("/", 1)[-1] or parsed.netloc
    return name[:28]


def _dot_quote(value: object) -> str:
    return json.dumps(str(value), ensure_ascii=False)


def _color_to_hex(value: object) -> str:
    try:
        from matplotlib.colors import to_hex

        return to_hex(value)
    except Exception:
        return "#999999"


def _write_dot(
    output: Path,
    graph: nx.DiGraph,
    *,
    labels: dict[object, str],
    node_colors: dict[object, object] | None = None,
    node_sizes: dict[object, int] | None = None,
    edge_labels: dict[tuple[object, object], object] | None = None,
    rankdir: str = "LR",
) -> None:
    """Write a Graphviz DOT file without requiring graphviz at runtime."""
    output.parent.mkdir(parents=True, exist_ok=True)
    node_ids = {node: f"n{index}" for index, node in enumerate(graph.nodes, start=1)}
    lines = [
        "digraph G {",
        f"  graph [rankdir={_dot_quote(rankdir)}, overlap=false, splines=true];",
        '  node [shape=box, style="rounded,filled", fontname="DejaVu Sans", fontsize=10, color="#333333"];',
        '  edge [fontname="DejaVu Sans", fontsize=8, color="#777777", arrowsize=0.6];',
    ]
    for node, node_id in node_ids.items():
        attrs = {
            "label": labels.get(node, str(node)),
            "tooltip": str(node),
            "fillcolor": _color_to_hex(node_colors.get(node, "#E8EEF7")) if node_colors else "#E8EEF7",
        }
        if node_sizes:
            attrs["width"] = round(max(0.8, min(3.2, node_sizes.get(node, 100) / 500)), 2)
        attr_text = ", ".join(f"{key}={_dot_quote(value)}" for key, value in attrs.items())
        lines.append(f"  {node_id} [{attr_text}];")
    for source, target in graph.edges:
        attrs = []
        if edge_labels and (source, target) in edge_labels:
            attrs.append(f"label={_dot_quote(edge_labels[(source, target)])}")
        attr_text = f" [{', '.join(attrs)}]" if attrs else ""
        lines.append(f"  {node_ids[source]} -> {node_ids[target]}{attr_text};")
    lines.append("}")
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _render_dot(dot_path: Path, png_path: Path, *, program: str) -> Path | None:
    """Render DOT to PNG using pygraphviz or graphviz if available.

    Graphviz is deliberately optional. When it is missing we keep the DOT file
    and the matplotlib/networkx PNGs, and log the reason instead of silently
    pretending a Graphviz render happened.
    """
    reasons: list[str] = []
    try:
        import pygraphviz as pgv  # type: ignore[import-not-found]

        graph = pgv.AGraph(str(dot_path))
        graph.layout(prog=program)
        graph.draw(str(png_path))
        return png_path
    except Exception as exc:
        reasons.append(f"pygraphviz: {type(exc).__name__}: {exc}")

    try:
        import graphviz  # type: ignore[import-not-found]

        source = graphviz.Source(dot_path.read_text(encoding="utf-8"), filename=png_path.with_suffix("").name, directory=str(png_path.parent), format="png", engine=program)
        rendered = Path(source.render(cleanup=True))
        if rendered != png_path and rendered.exists():
            rendered.replace(png_path)
        return png_path
    except Exception as exc:
        reasons.append(f"graphviz: {type(exc).__name__}: {exc}")

    LOGGER.warning(
        "GRAPHVIZ_RENDER_SKIPPED dot=%s target=%s program=%s dot_executable=%s reasons=%s",
        dot_path,
        png_path,
        program,
        shutil.which(program) or shutil.which("dot") or "",
        " | ".join(reasons),
    )
    return None


def _write_graphviz_status(graph_dir: Path) -> Path:
    output = graph_dir / "graphviz_status.txt"
    lines = [
        "Graphviz status",
        f"python_package_graphviz: {bool(importlib.util.find_spec('graphviz'))}",
        f"python_package_pygraphviz: {bool(importlib.util.find_spec('pygraphviz'))}",
        f"dot_executable: {shutil.which('dot') or ''}",
        f"sfdp_executable: {shutil.which('sfdp') or ''}",
        "",
        "DOT files are always written. PNG rendering via Graphviz is used when graphviz/pygraphviz and the Graphviz executables are available.",
        "If they are missing, the crawler keeps the networkx/matplotlib PNGs and logs GRAPHVIZ_RENDER_SKIPPED.",
    ]
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return output


def write_graphs(
    graph_dir: Path,
    *,
    pages: list[PageRecord],
    pdfs: list[PdfRecord],
    links: list[LinkRecord],
    errors: list[ErrorRecord],
    max_link_graph_nodes: int = 180,
) -> list[Path]:
    graph_dir.mkdir(parents=True, exist_ok=True)
    sns.set_theme(style="whitegrid", font="DejaVu Sans")
    generated: list[Path] = []
    generated.append(_write_graphviz_status(graph_dir))

    frames = build_stats_frames(pages=pages, pdfs=pdfs, links=links, errors=errors)
    depth_stats = frames["DepthStats"]
    if not depth_stats.empty:
        output = graph_dir / "depth_distribution.png"
        _save_bar(depth_stats, x="depth", y="pages", title="Pages by crawl depth", output=output, xlabel="depth", ylabel="pages")
        generated.append(output)

    section_stats = frames["SectionStats"].head(25)
    if not section_stats.empty:
        output = graph_dir / "section_pages.png"
        _save_bar(section_stats, x="section", y="pages", title="Pages by ISA section", output=output, xlabel="section", ylabel="pages")
        generated.append(output)

    pdf_stats = frames["PdfStats"].head(25)
    if not pdf_stats.empty:
        output = graph_dir / "pdfs_by_section.png"
        _save_bar(pdf_stats, x="source_section", y="pdf_references", title="PDF references by ISA section", output=output, xlabel="section", ylabel="PDF refs")
        generated.append(output)

    if errors:
        error_stats = frames["ErrorStats"]
        if not error_stats.empty:
            plot_frame = error_stats.copy()
            plot_frame["phase_type"] = plot_frame["phase"].astype(str) + " / " + plot_frame["exception_type"].astype(str)
            output = graph_dir / "errors_by_phase.png"
            _save_bar(plot_frame.head(25), x="phase_type", y="errors", title="Errors by phase", output=output, xlabel="phase / exception", ylabel="errors")
            generated.append(output)

    page_urls = {page.url for page in pages}
    graph = nx.DiGraph()
    pdf_count_by_page = {page.url: page.pdf_link_count for page in pages}
    section_by_page = {page.url: page.section for page in pages}
    title_by_page = {page.url: f"{page.section}/{_safe_graph_name(page.url)}" for page in pages}
    for page in pages:
        graph.add_node(page.url)
    for link in links:
        if link.category == "internal_page" and link.source_page_url in page_urls and link.url in page_urls:
            graph.add_edge(link.source_page_url, link.url)

    if graph.number_of_nodes() > 1 and graph.number_of_edges() > 0:
        score = {
            node: graph.degree(node) + pdf_count_by_page.get(node, 0) * 2
            for node in graph.nodes
        }
        selected = {node for node, _score in sorted(score.items(), key=lambda item: item[1], reverse=True)[:max_link_graph_nodes]}
        if pages:
            selected.add(pages[0].url)
        subgraph = graph.subgraph(selected).copy()
        output = graph_dir / "link_structure.png"
        plt.figure(figsize=(16, 12))
        pos = nx.spring_layout(subgraph, seed=42, k=0.35)
        sections = sorted({section_by_page.get(node, "") for node in subgraph.nodes})
        palette = sns.color_palette("tab20", n_colors=max(1, len(sections)))
        color_by_section = {section: palette[index] for index, section in enumerate(sections)}
        node_colors = [color_by_section.get(section_by_page.get(node, ""), "#999999") for node in subgraph.nodes]
        node_sizes = [80 + min(1200, graph.degree(node) * 18 + pdf_count_by_page.get(node, 0) * 25) for node in subgraph.nodes]
        nx.draw_networkx_edges(subgraph, pos, alpha=0.18, width=0.8, arrows=False)
        nx.draw_networkx_nodes(subgraph, pos, node_color=node_colors, node_size=node_sizes, alpha=0.86, linewidths=0.2, edgecolors="#333333")
        label_nodes = sorted(subgraph.nodes, key=lambda node: score.get(node, 0), reverse=True)[:35]
        labels = {node: title_by_page.get(node, _safe_graph_name(node))[:24] for node in label_nodes}
        nx.draw_networkx_labels(subgraph, pos, labels=labels, font_size=7)
        plt.title(f"Internal link structure, top {subgraph.number_of_nodes()} nodes")
        plt.axis("off")
        plt.tight_layout()
        plt.savefig(output, dpi=170)
        plt.close()
        generated.append(output)

        dot_output = graph_dir / "link_structure.dot"
        node_color_map = {node: color_by_section.get(section_by_page.get(node, ""), "#999999") for node in subgraph.nodes}
        node_size_map = {node: 80 + min(1200, graph.degree(node) * 18 + pdf_count_by_page.get(node, 0) * 25) for node in subgraph.nodes}
        _write_dot(
            dot_output,
            subgraph,
            labels={node: title_by_page.get(node, _safe_graph_name(node))[:40] for node in subgraph.nodes},
            node_colors=node_color_map,
            node_sizes=node_size_map,
        )
        generated.append(dot_output)
        rendered = _render_dot(dot_output, graph_dir / "link_structure_graphviz.png", program="sfdp")
        if rendered is not None:
            generated.append(rendered)

    page_to_section = {page.url: page.section for page in pages}
    section_graph = nx.DiGraph()
    edge_weights: dict[tuple[str, str], int] = {}
    for section in sorted(set(page_to_section.values())):
        section_graph.add_node(section)
    for link in links:
        if link.category != "internal_page":
            continue
        source_section = page_to_section.get(link.source_page_url)
        target_section = page_to_section.get(link.url)
        if not source_section or not target_section or source_section == target_section:
            continue
        edge = (source_section, target_section)
        edge_weights[edge] = edge_weights.get(edge, 0) + 1
    for (source_section, target_section), weight in edge_weights.items():
        section_graph.add_edge(source_section, target_section, weight=weight)
    if section_graph.number_of_nodes() > 0:
        section_dot = graph_dir / "section_links.dot"
        section_counts = {section: sum(1 for page in pages if page.section == section) for section in section_graph.nodes}
        section_pdf_counts = {section: sum(page.pdf_link_count for page in pages if page.section == section) for section in section_graph.nodes}
        section_labels = {
            section: f"{section}\npages={section_counts.get(section, 0)} pdf_refs={section_pdf_counts.get(section, 0)}"
            for section in section_graph.nodes
        }
        section_palette = sns.color_palette("tab20", n_colors=max(1, section_graph.number_of_nodes()))
        section_colors = {section: section_palette[index] for index, section in enumerate(section_graph.nodes)}
        _write_dot(
            section_dot,
            section_graph,
            labels=section_labels,
            node_colors=section_colors,
            node_sizes={section: 300 + section_counts.get(section, 0) * 25 for section in section_graph.nodes},
            edge_labels={(source, target): weight for (source, target), weight in edge_weights.items()},
            rankdir="TB",
        )
        generated.append(section_dot)
        rendered = _render_dot(section_dot, graph_dir / "section_links_graphviz.png", program="dot")
        if rendered is not None:
            generated.append(rendered)

    return generated
