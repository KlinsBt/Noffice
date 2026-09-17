# Dependency notice review

Updated September 12, 2026. The PDF integration raises the production inventory to 193 packages. Eight have missing/unknown full notices, including new `@pdf-lib/fontkit`; the complete transitive review is **unfinished**. `npm.cmd run notices` regenerated the actual inventory and shipped notice text in `.local/pagination-pdf-notices.log`.

`npm.cmd run notices` regenerates the shipped notices and inventory offline. `npm.cmd run notices:check` checks their exact generated contents without writing and exits nonzero for unresolved notices. `npm.cmd run notices:test` exercises absent/empty/directory notices, stale output, and tampered package/notice bindings. An existing file name alone is no longer counted as a readable notice. Passing this check establishes notice consistency, not complete provenance certification.

The reviewed supplement in `reviewed.json` binds package name/version, exact installed manifest bytes, lockfile integrity, upstream URL and notice hash. Any change requires renewed review. Do not infer a grant from another package's license or invent a copyright year.

| Package | Evidence and remaining action |
| --- | --- |
| saxes 5.0.1 | Resolved: the exact [v5.0.1 upstream LICENSE](https://github.com/lddubeau/saxes/blob/6ef6a275b20f19cb67808daf0d8f64e06aaf5b0e/LICENSE) and version manifest declare ISC. The complete notice includes inherited sax and historical MIT text; it is bundled verbatim. |
| @pdf-lib/fontkit 1.1.1 | Published manifest and README declare MIT but link a generic license template and omit a full attributed notice. Exact upstream notice and bundled component applicability remain under review. |
| woff2-encoder 2.0.0 | Published wrapper includes its MIT notice; independently review the bundled Google WOFF2/Brotli component notices before distribution. |
| binary 0.3.0 | Published manifest and README declare MIT, but no full notice is bundled. Original upstream repository returns 404; recover the applicable original notice/provenance. |
| buffers 0.1.1 | Published package has no license field. The original [commit preserved in a fork](https://github.com/TooTallNate/node-buffers/blob/1b745ee35d33eb166e15ef1866073a07c6d7de87/package.json) adds `MIT/X11`; full notice and applicability to the installed source still need review. No grant was synthesized. |
| chainsaw 0.1.0 | Published manifest declares `MIT/X11`; no full notice. Recover the original terms and attribution. |
| dingbat-to-unicode 1.0.1 | Published manifest declares BSD-2-Clause. The upstream repository is available but has no root license file; inspect the release tree and mapping-data provenance. |
| https 1.0.0 | Published archive installs only `package.json`, declaring ISC and its author; its advertised `index.js` is absent. Determine distribution applicability explicitly; do not silently drop it from the inventory. |
| is-reference 3.0.3 | Published manifest/README declare MIT. npm gitHead `8bb053129bfabe2f6a7d7ed050159d67ebe82829` has no full license file; recover original terms/attribution. |
| locate-character 3.0.0 | Published manifest/README declare MIT. npm gitHead `4f08a59ec248121f7002abd02ee7b94e8eda06bc` has no full license file; recover original terms/attribution. |

Local research receipts are in `.local/notice-research`. None of these unresolved cases were removed, relabelled as resolved from metadata alone, or excluded merely because bundling might tree-shake them. Font distribution, dual-license choices and external fixtures remain covered separately in `THIRD_PARTY_NOTICES.md`. Public distribution remains gated on completing the review and all other release requirements.
