
use corro::grid::Grid;
use corro::grid::GridBox;
use corro::grid::CellAddr;
use corro::grid::MARGIN_COLS;
use corro::grid::HEADER_ROWS;

fn main(){
    let mut g=GridBox::from(Grid::new(2,3));
    let mc=g.main_cols();
    let right_a=(MARGIN_COLS+mc) as u32;
    let addr=CellAddr::Header { row:(HEADER_ROWS-1) as u32, col:right_a };
    g.set(&addr, "=B".to_string());
    println!("set {} to =B", right_a);
    println!("grid.get header: {:?}", g.get(&addr));
    let main_addr=CellAddr::Main{row:0, col:(mc-1) as u32};
    println!("templated_formula main: {:?}", corro::formula::templated_formula(&g, &main_addr));
    
    // print header cells around margins
    for c in [MARGIN_COLS-1, MARGIN_COLS, MARGIN_COLS+mc, MARGIN_COLS+mc+1]:
        a=CellAddr::Header{row:(HEADER_ROWS-1) as u32, col:c as u32}
        print(f"col {c}: {g.get(&a)}
")
}
