#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum BalanceDirection {
    PosToNeg,
    NegToPos,
}

pub fn parse_amount_cents(raw: &str) -> Option<i64> {
    let t = raw.trim();
    if t.is_empty() {
        return None;
    }

    let mut s = t;
    let negative = match s.chars().next()? {
        '+' => {
            s = &s[1..];
            false
        }
        '-' => {
            s = &s[1..];
            true
        }
        _ => false,
    };

    let s = s.replace(',', "");
    let (whole, frac) = match s.split_once('.') {
        Some((whole, frac)) => (whole, frac),
        None => (s.as_str(), ""),
    };
    if whole.is_empty() || !whole.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    if !frac.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }

    let mut cents = whole.parse::<i64>().ok()?.saturating_mul(100);
    let frac_cents = match frac.len() {
        0 => 0,
        1 => frac.parse::<i64>().ok()?.saturating_mul(10),
        _ => frac.get(0..2)?.parse::<i64>().ok()?,
    };
    cents = cents.saturating_add(frac_cents);
    Some(if negative { -cents } else { cents })
}

pub fn format_amount_cents(cents: i64) -> String {
    let sign = if cents < 0 { "-" } else { "" };
    let abs = cents.abs();
    format!("{sign}{}.{:02}", abs / 100, abs % 100)
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceItem {
    pub id: usize,
    pub amount_cents: i64,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceGroup {
    pub item_ids: Vec<usize>,
    pub total_cents: i64,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceSolution {
    pub groups: Vec<BalanceGroup>,
    pub leftovers: Vec<usize>,
}

pub fn solve_balance_groups(items: &[BalanceItem], direction: BalanceDirection) -> BalanceSolution {
    let mut supply: Vec<BalanceItem> = items
        .iter()
        .filter(|item| match direction {
            BalanceDirection::PosToNeg => item.amount_cents > 0,
            BalanceDirection::NegToPos => item.amount_cents < 0,
        })
        .cloned()
        .collect();
    let mut demand: Vec<BalanceItem> = items
        .iter()
        .filter(|item| match direction {
            BalanceDirection::PosToNeg => item.amount_cents < 0,
            BalanceDirection::NegToPos => item.amount_cents > 0,
        })
        .cloned()
        .collect();

    supply.sort_by_key(|r| (r.id, r.amount_cents.abs()));
    demand.sort_by_key(|r| (r.id, r.amount_cents.abs()));

    let mut used = vec![false; demand.len()];
    let mut groups = Vec::new();

    for anchor in supply {
        let target = anchor.amount_cents.abs();
        if let Some(indices) = choose_subset(&demand, &used, target) {
            for &idx in &indices {
                used[idx] = true;
            }
            let total: i64 = indices.iter().map(|&idx| demand[idx].amount_cents).sum();
            groups.push(BalanceGroup {
                item_ids: std::iter::once(anchor.id)
                    .chain(indices.into_iter().map(|idx| demand[idx].id))
                    .collect(),
                total_cents: anchor.amount_cents + total,
            });
        }
    }

    let mut marked = vec![false; items.len()];
    for group in &groups {
        for &item_id in &group.item_ids {
            if item_id < marked.len() {
                marked[item_id] = true;
            }
        }
    }
    let leftovers = (0..items.len()).filter(|idx| !marked[*idx]).collect();

    BalanceSolution { groups, leftovers }
}

fn choose_subset(items: &[BalanceItem], used: &[bool], target: i64) -> Option<Vec<usize>> {
    fn rec(
        items: &[BalanceItem],
        used: &[bool],
        target: i64,
        start: usize,
        current: &mut Vec<usize>,
    ) -> Option<Vec<usize>> {
        if target == 0 {
            return Some(current.clone());
        }
        if target < 0 {
            return None;
        }

        for idx in start..items.len() {
            if used[idx] {
                continue;
            }
            let value = items[idx].amount_cents.abs();
            if value > target {
                continue;
            }
            current.push(idx);
            if let Some(found) = rec(items, used, target - value, idx + 1, current) {
                return Some(found);
            }
            current.pop();
        }
        None
    }

    rec(items, used, target, 0, &mut Vec::new())
}
