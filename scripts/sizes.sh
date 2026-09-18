#for f in pancurses pancurses,gui ratatui ratatui,gui ratatui,pancurses ratatui,pancurses,gui
for f in gui
do
	cargo build --release --no-default-features --features $f
	echo `stat %s target/release/corro` $f | tee -a size.log
done
cat size.log | sed 's/.*Size: //;s/Blocks.*+0800 //' | sort
