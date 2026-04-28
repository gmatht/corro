package main

import (
	"bufio"
	"context"
	"errors"
	"flag"
	"fmt"
	"io"
	"math"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
	"strings"
	"syscall"
	"time"
)

type movieEntry struct {
	FilePath string
	Title    string
	Desc     string
}

type parseState struct {
	filePath string
	title    string
	desc     string
	start    int
}

func (p *parseState) empty() bool {
	return p.filePath == "" && p.title == "" && p.desc == "" && p.start == 0
}

func (p *parseState) setField(key string, value string, lineNo int) error {
	switch key {
	case "FILE":
		if p.filePath != "" {
			return fmt.Errorf("line %d: duplicate FILE field", lineNo)
		}
		p.filePath = value
	case "TITLE":
		if p.title != "" {
			return fmt.Errorf("line %d: duplicate TITLE field", lineNo)
		}
		p.title = value
	case "DESC":
		if p.desc != "" {
			return fmt.Errorf("line %d: duplicate DESC field", lineNo)
		}
		p.desc = value
	default:
		return fmt.Errorf("line %d: unknown key %q", lineNo, key)
	}
	if p.start == 0 {
		p.start = lineNo
	}
	return nil
}

func resolveEntryPath(rootDir string, rawPath string) string {
	if filepath.IsAbs(rawPath) {
		return filepath.Clean(rawPath)
	}

	cwd, err := os.Getwd()
	if err == nil {
		candidate := filepath.Clean(filepath.Join(cwd, rawPath))
		if _, statErr := os.Stat(candidate); statErr == nil {
			return candidate
		}
	}

	return filepath.Clean(filepath.Join(rootDir, rawPath))
}

func (p *parseState) finish(entries *[]movieEntry, rootDir string, lineNo int) error {
	if p.empty() {
		return nil
	}

	if p.filePath == "" || p.title == "" || p.desc == "" {
		return fmt.Errorf("line %d: incomplete block (need FILE, TITLE, DESC)", lineNo)
	}

	if !strings.EqualFold(filepath.Ext(p.filePath), ".corro") {
		return fmt.Errorf("line %d: FILE must end with .corro: %s", p.start, p.filePath)
	}

	resolved := resolveEntryPath(rootDir, p.filePath)

	info, err := os.Stat(resolved)
	if err != nil {
		return fmt.Errorf("line %d: FILE does not exist: %s", p.start, p.filePath)
	}
	if info.IsDir() {
		return fmt.Errorf("line %d: FILE points to directory: %s", p.start, p.filePath)
	}

	*entries = append(*entries, movieEntry{
		FilePath: resolved,
		Title:    p.title,
		Desc:     p.desc,
	})

	*p = parseState{}
	return nil
}

func parseMovieScript(scriptPath string) ([]movieEntry, error) {
	f, err := os.Open(scriptPath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	rootDir := filepath.Dir(scriptPath)
	scanner := bufio.NewScanner(f)
	// Raise default token limit for longer description lines.
	scanner.Buffer(make([]byte, 1024), 1024*1024)

	entries := make([]movieEntry, 0, 8)
	var state parseState
	lineNo := 0

	for scanner.Scan() {
		lineNo++
		raw := scanner.Text()
		line := strings.TrimSpace(raw)

		if line == "" {
			if err := state.finish(&entries, rootDir, lineNo); err != nil {
				return nil, err
			}
			continue
		}

		parts := strings.SplitN(line, ":", 2)
		if len(parts) != 2 {
			return nil, fmt.Errorf("line %d: expected KEY: value format", lineNo)
		}

		key := strings.ToUpper(strings.TrimSpace(parts[0]))
		value := strings.TrimSpace(parts[1])
		if value == "" {
			return nil, fmt.Errorf("line %d: empty value for %s", lineNo, key)
		}

		if err := state.setField(key, value, lineNo); err != nil {
			return nil, err
		}
	}
	if err := scanner.Err(); err != nil {
		return nil, err
	}
	if err := state.finish(&entries, rootDir, lineNo+1); err != nil {
		return nil, err
	}
	if len(entries) == 0 {
		return nil, errors.New("movie script has no entries")
	}

	return entries, nil
}

func ansiBlueBackground(w io.Writer) {
	_, _ = fmt.Fprint(w, "\x1b[44m\x1b[2J\x1b[H")
}

func ansiResetAndClear(w io.Writer) {
	_, _ = fmt.Fprint(w, "\x1b[0m\x1b[2J\x1b[H")
}

func typeText(ctx context.Context, w io.Writer, text string, cps float64) error {
	if cps <= 0 {
		cps = 12
	}
	perChar := time.Duration(float64(time.Second) / cps)
	if perChar < 5*time.Millisecond {
		perChar = 5 * time.Millisecond
	}

	for _, r := range text {
		select {
		case <-ctx.Done():
			return ctx.Err()
		default:
		}

		_, err := fmt.Fprint(w, string(r))
		if err != nil {
			return err
		}
		time.Sleep(perChar)
	}
	_, _ = fmt.Fprintln(w)
	return nil
}

func commandExists(name string) bool {
	_, err := exec.LookPath(name)
	return err == nil
}

func runCommandStreaming(ctx context.Context, command string, args ...string) error {
	cmd := exec.CommandContext(ctx, command, args...)
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	cmd.Stdin = os.Stdin
	if err := cmd.Run(); err != nil {
		return fmt.Errorf("%s %v: %w", command, args, err)
	}
	return nil
}

func printTitleWithCowsay(ctx context.Context, title string, forceWSL bool) {
	switch {
	case !forceWSL && commandExists("cowsay"):
		if err := runCommandStreaming(ctx, "cowsay", title); err == nil {
			return
		}
	case commandExists("wsl"):
		if err := runCommandStreaming(ctx, "wsl", "cowsay", title); err == nil {
			return
		}
	}

	// Graceful fallback if cowsay isn't available.
	fmt.Printf("=== %s ===\n", title)
}

func runCorroMovie(ctx context.Context, corroBin string, filePath string, forceWSL bool) error {
	if forceWSL {
		return runCommandStreaming(ctx, "wsl", corroBin, "--movie", filePath)
	}
	return runCommandStreaming(ctx, corroBin, "--movie", filePath)
}

func main() {
	scriptFlag := flag.String("script", "scripts/movie_demo.txt", "movie script file path")
	pauseMsFlag := flag.Int("pause-ms", 2000, "pause after description (ms)")
	cpsFlag := flag.Float64("cps", 12.0, "description typing speed (chars per second)")
	corroBinFlag := flag.String("corro-bin", "corro", "corro executable name/path")
	wslFlag := flag.Bool("wsl", false, "force invocation through wsl for cowsay/corro")
	flag.Parse()

	if *pauseMsFlag < 0 {
		fmt.Fprintln(os.Stderr, "pause-ms must be >= 0")
		os.Exit(2)
	}
	if *cpsFlag <= 0 || math.IsNaN(*cpsFlag) || math.IsInf(*cpsFlag, 0) {
		fmt.Fprintln(os.Stderr, "cps must be a positive finite number")
		os.Exit(2)
	}

	scriptPath := *scriptFlag
	if !filepath.IsAbs(scriptPath) {
		cwd, err := os.Getwd()
		if err != nil {
			fmt.Fprintf(os.Stderr, "failed to get working directory: %v\n", err)
			os.Exit(1)
		}
		scriptPath = filepath.Join(cwd, scriptPath)
	}
	scriptPath = filepath.Clean(scriptPath)

	entries, err := parseMovieScript(scriptPath)
	if err != nil {
		fmt.Fprintf(os.Stderr, "movie script parse error: %v\n", err)
		os.Exit(1)
	}

	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt, syscall.SIGTERM)
	defer stop()

	defer ansiResetAndClear(os.Stdout)

	pause := time.Duration(*pauseMsFlag) * time.Millisecond
	total := len(entries)

	for idx, entry := range entries {
		select {
		case <-ctx.Done():
			fmt.Fprintln(os.Stderr, "interrupted")
			os.Exit(130)
		default:
		}

		fmt.Printf("\n[%d/%d] %s\n", idx+1, total, entry.FilePath)
		ansiBlueBackground(os.Stdout)
		printTitleWithCowsay(ctx, entry.Title, *wslFlag)

		if err := typeText(ctx, os.Stdout, entry.Desc, *cpsFlag); err != nil {
			fmt.Fprintf(os.Stderr, "typing failed: %v\n", err)
			os.Exit(1)
		}

		time.Sleep(pause)
		ansiResetAndClear(os.Stdout)

		if err := runCorroMovie(ctx, *corroBinFlag, entry.FilePath, *wslFlag); err != nil {
			fmt.Fprintf(os.Stderr, "failed running corro --movie for %s: %v\n", entry.FilePath, err)
			os.Exit(1)
		}
	}
}
