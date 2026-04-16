Grid Abstraction Plan

Goal
Abstract away the concrete Grid implementation so most of the codebase does not assume a HashMap<(u32,u32)> representation. The goal is incremental: introduce an interface, adapt the majority of usages, keep behavior identical, and make further optimizations possible.

Principles
- Make the smallest correct changes first.
- Keep the public API surface minimal and explicit.
- Preserve current behavior and tests during the migration.

Milestones
1. Introduce Grid trait
  - Create a trait (GridImpl) describing required operations used across the codebase: get, set, remove, main_rows/main_cols sizing, iterate non-empty cells in ranges, volatile_seed accessor, and any allocation/growth operations (set_main_size, ensure_main_size, clear_range, etc.).
  - Add a thin wrapper type (Grid) currently implemented by the existing in-memory HashMap-backed struct (call it HashGrid or GridMap). Keep the current Grid struct name if it already exists, but make it delegate to a boxed GridImpl so most call sites use Grid methods rather than the underlying map.

2. Convert internal consumers to use Grid API
  - Search for direct hashmap accessors (e.g., grid.map or grid.inner) and replace them with Grid methods.
  - Update grid creation sites to instantiate the wrapper backed by HashGrid.
  - Keep a minimal set of internal helper functions for iterating cells to avoid code duplication.

3. Add adapters/compat shims
  - Where external modules depend on specific Grid internals (rare), add adapter functions on Grid to expose only necessary functionality (for example, an iterator over non-empty cells). Avoid exposing internal HashMap type.

4. Tests and validation
  - Run full test suite after each milestone.
  - Add unit tests for Grid trait and the HashGrid implementation ensuring identical behavior (sizes, set/get semantics, persistence across ops like move_cols/move_rows, volatile_seed behavior).

5. Progressively remove accidental leaks
  - After all call sites use Grid methods, remove any remaining public fields exposing internals, then make underlying storage private.

6. Optional: introduce alternative implementations
  - Add an in-memory Vec-backed dense grid implementation for benchmarks and memory-usage comparison.
  - Provide a persistent disk-backed variant for future features if desired.

Detailed Steps
1. Add trait and wrapper
  - Add src/grid/trait.rs (or augment src/grid/mod.rs) with the GridImpl trait describing required methods:
    - fn get(&self, addr: &CellAddr) -> Option<&str> or Option<String>
    - fn set(&mut self, addr: &CellAddr, value: String)
    - fn remove(&mut self, addr: &CellAddr)
    - fn main_rows(&self) -> usize
    - fn main_cols(&self) -> usize
    - fn set_main_size(&mut self, r: usize, c: usize)
    - fn total_cols(&self) -> usize
    - fn iter_nonempty(&self) -> Box<dyn Iterator<Item=(CellAddr, String)>>
    - volatile_seed getters/setters
  - Add Grid wrapper that holds Box<dyn GridImpl> and implements the same methods delegating to the inner impl.

2. Implement GridMap (current behavior)
  - Move current HashMap-backed logic into GridMap implementing GridImpl.
  - Ensure GridMap's methods mirror current Grid behavior exactly.

3. Replace direct uses
  - Use grep to find .map or direct hash accesses and replace with Grid::get/set/iter calls.
  - Adjust imports and type signatures where functions currently accept concrete HashMap type; replace with &Grid or &mut Grid.

4. Tighten encapsulation
  - Make old public fields private.
  - Ensure serialization and log replay use Grid wrapper methods.

5. Add tests and ensure parity
  - Add tests that exercise serialization, move operations, formatting, and behavior for header/footer addressing.
  - Fix any test breakages incrementally.

Time estimate
- Milestone 1-2: 2-4 hours
- Milestone 3-4: 2-6 hours depending on number of call sites
- Tests & polishing: 1-3 hours

Risks
- Large surface area: many modules interact with the grid; careful incremental replacements and frequent test runs will mitigate risk.
- Performance: dynamic dispatch may add overhead; if it matters, add a generic wrapper or use enum-based dispatch after correctness.

Next immediate actions I will take (if you want me to proceed now)
1. Add the GridImpl trait and Grid wrapper with the HashMap-backed implementation.
2. Replace simple call sites (get/set/set_main_size) to use the wrapper.
3. Run the test suite and report failures.
