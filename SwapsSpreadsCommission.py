# Quick script to get Spreads, Swaps and commission from broker
# Exported to multi-worksheet xlsx file in working directory

import MetaTrader5 as mt5
import pandas as pd
from datetime import datetime
import time

# MT5 Connection Settings
MT5_PATH = r"C:\Program Files\broker\terminal64.exe"
ACCOUNT = 1234567890
PASSWORD = "***********"
SERVER = "Broker-Server"

# Currency pairs
ALL_CCY_PAIRS = [
    # Major Pairs
    'EURUSD', 'GBPUSD', 'USDJPY', 'AUDUSD', 'USDCAD', 'USDCHF', 'NZDUSD',

    # EUR Crosses
    'EURJPY', 'EURCAD', 'EURCHF', 'EURGBP', 'EURAUD', 'EURNZD', 'EURSGD',
    'EURDKK', 'EURHKD', 'EURNOK', 'EURPLN', 'EURSEK', 'EURTRY', 'EURZAR',

    # GBP Crosses
    'GBPCHF', 'GBPJPY', 'GBPAUD', 'GBPCAD', 'GBPNZD', 'GBPDKK', 'GBPNOK',
    'GBPSEK', 'GBPSGD', 'GBPTRY',

    # AUD Crosses
    'AUDNZD', 'AUDCAD', 'AUDCHF', 'AUDJPY', 'AUDSGD',

    # CAD Crosses
    'CADCHF', 'CADJPY',

    # CHF Crosses
    'CHFJPY', 'CHFSGD',

    # NZD Crosses
    'NZDCAD', 'NZDCHF', 'NZDJPY', 'NZDSGD',

    # JPY Crosses
    'NOKJPY', 'SEKJPY', 'SGDJPY',

    # Other USD Crosses
    'USDSGD', 'USDDKK', 'USDHKD', 'USDNOK', 'USDPLN', 'USDSEK', 'USDTRY',
    'USDZAR', 'USDCNH', 'USDCZK', 'USDHUF', 'USDMXN', 'USDTHB',

    # Other Crosses
    'NOKSEK',

    # Commodities
    'XAUUSD', 'XAGUSD', 'XAUEUR', 'XPDUSD',

    # Indices
    'US30'
]


class SwapSpreadCommissionMonitor:
    def __init__(self, path, account, password, server):
        self.path = path
        self.account = account
        self.password = password
        self.server = server
        self.connected = False

    def connect(self):
        """Establish connection to MT5"""
        if not mt5.initialize(path=self.path):
            print("MT5 initialization failed")
            return False

        authorized = mt5.login(
            login=self.account,
            password=self.password,
            server=self.server
        )

        if authorized:
            print(f"Connected to account {self.account}")
            self.connected = True
            return True
        else:
            print(f"Login failed. Error: {mt5.last_error()}")
            return False

    def safe_getattr(self, obj, attr, default=None):
        """Safely get attribute from object"""
        try:
            return getattr(obj, attr, default)
        except:
            return default

    def get_trading_costs_for_symbol(self, symbol):
        """Get detailed swap rates, spread, and commission for a specific symbol"""
        try:
            # Select the symbol
            selected = mt5.symbol_select(symbol, True)
            if not selected:
                return None

            time.sleep(0.05)  # Small delay

            symbol_info = mt5.symbol_info(symbol)
            if symbol_info is None:
                return None

            # Get current tick
            tick = mt5.symbol_info_tick(symbol)

            # Interpret swap mode safely
            swap_modes = {
                0: "Points",
                1: "Currency",
                2: "Interest",
                3: "Margin currency"
            }

            swap_mode_value = self.safe_getattr(symbol_info, 'swap_mode', -1)
            swap_mode = swap_modes.get(swap_mode_value, f"Mode {swap_mode_value}")

            # Get swap values safely
            swap_long = self.safe_getattr(symbol_info, 'swap_long', 0)
            swap_short = self.safe_getattr(symbol_info, 'swap_short', 0)
            point = self.safe_getattr(symbol_info, 'point', 0.00001)
            trade_contract_size = self.safe_getattr(symbol_info, 'trade_contract_size', 100000)

            # GET COMMISSION DATA
            commission = self.safe_getattr(symbol_info, 'commission', 0)
            commission_type = self.safe_getattr(symbol_info, 'commission_type', 0)
            commission_charge = self.safe_getattr(symbol_info, 'commission_charge', 0)  # Charge mode

            # Interpret commission type
            commission_types = {
                0: "Per lot (volume)",
                1: "Per money",
                2: "Per deal (one-way)",
                3: "Per deal (round-turn)"
            }
            commission_type_desc = commission_types.get(commission_type, f"Type {commission_type}")

            # Interpret commission charge mode
            charge_modes = {
                0: "Per trade",
                1: "Per 0.01 lot",
                2: "Per 0.1 lot",
                3: "Per 1 lot",
                4: "Per 10 lots",
                5: "Per 100 lots",
                6: "Per deal",
                7: "Per trade (profit)"
            }
            commission_charge_desc = charge_modes.get(commission_charge, f"Mode {commission_charge}")

            # Calculate approximate swap value
            swap_long_value = None
            swap_short_value = None

            if swap_mode_value == 0 and point > 0:  # Points
                # Calculate value in account currency (usually USD)
                swap_long_value = swap_long * point * trade_contract_size
                swap_short_value = swap_short * point * trade_contract_size
            elif swap_mode_value == 1:  # Currency
                swap_long_value = swap_long
                swap_short_value = swap_short
            else:
                # For other modes or if calculation fails
                swap_long_value = swap_long
                swap_short_value = swap_short

            # Calculate commission per standard lot (100,000 units)
            commission_per_lot = None
            if commission_type == 0:  # Per lot (most common)
                commission_per_lot = commission  # Already in account currency per lot
            elif commission_type == 1:  # Per money (per $1,000,000)
                commission_per_lot = commission * trade_contract_size / 1000000
            elif commission_type in [2, 3]:  # Per deal
                commission_per_lot = commission
            else:
                commission_per_lot = commission

            # Adjust commission based on charge mode
            if commission_charge == 1:  # Per 0.01 lot
                commission_per_lot = commission_per_lot * 100
            elif commission_charge == 2:  # Per 0.1 lot
                commission_per_lot = commission_per_lot * 10
            elif commission_charge == 4:  # Per 10 lots
                commission_per_lot = commission_per_lot / 10
            elif commission_charge == 5:  # Per 100 lots
                commission_per_lot = commission_per_lot / 100

            # Calculate commission per day (amortized over 30 days for comparison)
            commission_daily = commission_per_lot / 30 if commission_per_lot else 0

            # Calculate total daily cost (swap + commission)
            total_cost_long = None
            total_cost_short = None
            if isinstance(swap_long_value, (int, float)) and isinstance(commission_daily, (int, float)):
                total_cost_long = swap_long_value + commission_daily
                total_cost_short = swap_short_value + commission_daily

            # Get triple swap day info
            triple_swap = self.safe_getattr(symbol_info, 'swap_rollover3days', False)

            # Get bid/ask
            bid = tick.bid if tick and hasattr(tick, 'bid') else 0
            ask = tick.ask if tick and hasattr(tick, 'ask') else 0

            # Calculate spread in pips
            spread_pips = 0
            spread_cost = 0
            if bid > 0 and ask > 0 and point > 0:
                spread_pips = round((ask - bid) / point, 1)
                # Calculate spread cost in account currency
                spread_cost = (ask - bid) * trade_contract_size

            # Calculate total entry cost (spread + commission)
            total_entry_cost = spread_cost + (
                commission_per_lot if commission_type in [2, 3] else commission_per_lot * 2)

            # Calculate pip value (approximate)
            pip_value = None
            if point > 0:
                pip_value = point * trade_contract_size

            return {
                'Symbol': symbol,
                'Bid': round(bid, 5) if bid > 0 else 'N/A',
                'Ask': round(ask, 5) if ask > 0 else 'N/A',
                'Spread (pips)': spread_pips if spread_pips > 0 else 'N/A',
                'Spread Cost ($/lot)': round(spread_cost, 2) if spread_cost > 0 else 'N/A',
                'Pip Value ($)': round(pip_value, 2) if pip_value else 'N/A',

                # Swap data
                'Swap Long Rate': swap_long,
                'Swap Short Rate': swap_short,
                'Swap Mode': swap_mode,
                'Swap Long ($/lot/day)': round(swap_long_value, 2) if isinstance(swap_long_value,
                                                                                 (int, float)) else 'N/A',
                'Swap Short ($/lot/day)': round(swap_short_value, 2) if isinstance(swap_short_value,
                                                                                   (int, float)) else 'N/A',
                'Triple Swap': 'Yes' if triple_swap else 'No',

                # Commission data
                'Commission': commission,
                'Commission Type': commission_type_desc,
                'Commission Charge': commission_charge_desc,
                'Commission ($/lot)': round(commission_per_lot, 2) if isinstance(commission_per_lot,
                                                                                 (int, float)) else 'N/A',
                'Commission Daily ($/lot/day)': round(commission_daily, 2) if isinstance(commission_daily,
                                                                                         (int, float)) else 'N/A',

                # Total costs
                'Total Cost Long ($/lot/day)': round(total_cost_long, 2) if isinstance(total_cost_long,
                                                                                       (int, float)) else 'N/A',
                'Total Cost Short ($/lot/day)': round(total_cost_short, 2) if isinstance(total_cost_short,
                                                                                         (int, float)) else 'N/A',
                'Total Entry Cost ($/lot)': round(total_entry_cost, 2) if isinstance(total_entry_cost,
                                                                                     (int, float)) else 'N/A',

                # Technical data
                'Point': point,
                'Contract Size': trade_contract_size,
                'Digits': self.safe_getattr(symbol_info, 'digits', 5)
            }

        except Exception as e:
            print(f"Error getting data for {symbol}: {str(e)[:50]}")
            return None

    def get_all_trading_costs(self):
        """Get swap rates, spreads, and commissions for all symbols"""
        if not self.connected:
            print("Not connected to MT5")
            return None

        print(f"\n{'=' * 80}")
        print(f"FETCHING TRADING COSTS - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Total symbols to check: {len(ALL_CCY_PAIRS)}")
        print(f"{'=' * 80}\n")

        cost_data = []
        success_count = 0
        fail_count = 0

        for i, symbol in enumerate(ALL_CCY_PAIRS, 1):
            print(f"Processing {i:3d}/{len(ALL_CCY_PAIRS)}: {symbol:8s}", end="\r")

            cost_info = self.get_trading_costs_for_symbol(symbol)
            if cost_info:
                cost_data.append(cost_info)
                success_count += 1
            else:
                fail_count += 1

            time.sleep(0.03)  # Small delay

        print(f"\n\nSuccessfully retrieved: {success_count} symbols")
        print(f"Failed to retrieve: {fail_count} symbols")
        print("=" * 80)

        return pd.DataFrame(cost_data) if cost_data else None

    def display_summary(self, df):
        """Display swap rates, spreads, and commissions in a formatted table"""
        if df is None or df.empty:
            print("No trading cost data available")
            return

        # Sort by symbol name
        df = df.sort_values('Symbol')

        print("\n" + "=" * 140)
        print("TRADING COSTS SUMMARY - ALL CURRENCY PAIRS")
        print("=" * 140)
        print("\nNote: All values per standard lot")
        print("      Swap: Negative = PAY swap, Positive = EARN swap")
        print("      Total Cost = Swap + Commission (amortized over 30 days)")
        print("      Entry Cost = Spread + Commission (one round turn)")
        print("=" * 140)

        # Group by type
        print("\nMAJOR PAIRS:")
        print("-" * 140)
        majors = ['EURUSD', 'GBPUSD', 'USDJPY', 'AUDUSD', 'USDCAD', 'USDCHF', 'NZDUSD']
        self._display_group(df, majors)

        print("\nEUR CROSSES:")
        print("-" * 140)
        eur_crosses = [p for p in df['Symbol'] if p.startswith('EUR') and p != 'EURUSD']
        self._display_group(df, eur_crosses)

        print("\nGBP CROSSES:")
        print("-" * 140)
        gbp_crosses = [p for p in df['Symbol'] if p.startswith('GBP') and p != 'GBPUSD']
        self._display_group(df, gbp_crosses)

        print("\nCOMMODITIES & INDICES:")
        print("-" * 140)
        commodities = [p for p in df['Symbol'] if p.startswith('X') or p in ['US30']]
        self._display_group(df, commodities)

        print("\nOTHER PAIRS:")
        print("-" * 140)
        other_pairs = [p for p in df['Symbol'] if p not in majors + eur_crosses + gbp_crosses + commodities]
        self._display_group(df, other_pairs)

        print(f"\n{'=' * 140}")
        print(f"Total Symbols with Data: {len(df)}")
        print(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'=' * 140}")

    def _display_group(self, df, symbols):
        """Display a group of symbols"""
        if not symbols:
            return

        group_df = df[df['Symbol'].isin(symbols)]
        if group_df.empty:
            print("No data for this group")
            return

        # Create formatted output
        output = f"{'Symbol':<8} {'Spread':<8} {'Swap Long':<12} {'Swap Short':<12} {'Commission':<12} {'Total Long':<14} {'Total Short':<14}\n"
        output += "-" * 90 + "\n"

        for _, row in group_df.iterrows():
            swap_long = row['Swap Long ($/lot/day)']
            swap_short = row['Swap Short ($/lot/day)']
            commission = row['Commission ($/lot)']
            total_long = row['Total Cost Long ($/lot/day)']
            total_short = row['Total Cost Short ($/lot/day)']

            long_str = f"${swap_long:+.2f}" if isinstance(swap_long, (int, float)) else str(swap_long)
            short_str = f"${swap_short:+.2f}" if isinstance(swap_short, (int, float)) else str(swap_short)
            comm_str = f"${commission:.2f}" if isinstance(commission, (int, float)) else str(commission)
            total_long_str = f"${total_long:+.2f}" if isinstance(total_long, (int, float)) else str(total_long)
            total_short_str = f"${total_short:+.2f}" if isinstance(total_short, (int, float)) else str(total_short)

            output += (f"{row['Symbol']:<8} "
                       f"{row['Spread (pips)']:<8} "
                       f"{long_str:<12} "
                       f"{short_str:<12} "
                       f"{comm_str:<12} "
                       f"{total_long_str:<14} "
                       f"{total_short_str:<14}\n")

        print(output)

    def analyze_trading_costs(self, df):
        """Analyze swap, spread, and commission data for trading opportunities"""
        if df is None or df.empty:
            return

        print("\n" + "=" * 90)
        print("COMPREHENSIVE TRADING COST ANALYSIS")
        print("=" * 90)

        # Filter only symbols with valid data
        valid_df = df[
            df['Swap Long ($/lot/day)'].apply(lambda x: isinstance(x, (int, float))) &
            df['Commission ($/lot)'].apply(lambda x: isinstance(x, (int, float))) &
            df['Spread Cost ($/lot)'].apply(lambda x: isinstance(x, (int, float)))
            ].copy()

        if valid_df.empty:
            print("No valid trading cost data for analysis")
            return

        # Ensure total cost columns exist
        if 'Total Cost Long ($/lot/day)' not in valid_df.columns:
            valid_df['Total Cost Long ($/lot/day)'] = valid_df['Swap Long ($/lot/day)'] + (
                        valid_df['Commission ($/lot)'] / 30)
            valid_df['Total Cost Short ($/lot/day)'] = valid_df['Swap Short ($/lot/day)'] + (
                        valid_df['Commission ($/lot)'] / 30)

        if 'Total Entry Cost ($/lot)' not in valid_df.columns:
            valid_df['Total Entry Cost ($/lot)'] = valid_df['Spread Cost ($/lot)'] + valid_df['Commission ($/lot)']

        # 1. BEST CARRY TRADES (considering all costs)
        print("\n1. BEST CARRY TRADES (Highest Net Positive):")
        print("-" * 90)

        # For LONG positions
        long_best = valid_df[valid_df['Total Cost Long ($/lot/day)'] > 0].copy()
        if not long_best.empty:
            long_best = long_best.sort_values('Total Cost Long ($/lot/day)', ascending=False)
            print("\nTop 5 LONG positions (earn swap after commission):")
            for i, (_, row) in enumerate(long_best.head(5).iterrows(), 1):
                print(f"{i:2d}. {row['Symbol']:8s}: "
                      f"Net Earn ${row['Total Cost Long ($/lot/day)']:+.2f}/day "
                      f"(Swap: ${row['Swap Long ($/lot/day)']:+.2f}, "
                      f"Comm: ${row['Commission ($/lot)']:.2f}, "
                      f"Spread: {row['Spread (pips)']} pips)")

        # For SHORT positions
        short_best = valid_df[valid_df['Total Cost Short ($/lot/day)'] > 0].copy()
        if not short_best.empty:
            short_best = short_best.sort_values('Total Cost Short ($/lot/day)', ascending=False)
            print("\nTop 5 SHORT positions (earn swap after commission):")
            for i, (_, row) in enumerate(short_best.head(5).iterrows(), 1):
                print(f"{i:2d}. {row['Symbol']:8s}: "
                      f"Net Earn ${row['Total Cost Short ($/lot/day)']:+.2f}/day "
                      f"(Swap: ${row['Swap Short ($/lot/day)']:+.2f}, "
                      f"Comm: ${row['Commission ($/lot)']:.2f}, "
                      f"Spread: {row['Spread (pips)']} pips)")

        # 2. LOWEST TOTAL HOLDING COST (for swing trading)
        print("\n\n2. LOWEST HOLDING COST PAIRS (Best for Swing Trading):")
        print("-" * 90)

        valid_df['Avg Holding Cost'] = (
                                               abs(valid_df['Total Cost Long ($/lot/day)']) +
                                               abs(valid_df['Total Cost Short ($/lot/day)'])
                                       ) / 2

        low_holding_cost = valid_df.sort_values('Avg Holding Cost').head(10)

        print("\nPairs with lowest daily holding costs:")
        for i, (_, row) in enumerate(low_holding_cost.iterrows(), 1):
            print(f"{i:2d}. {row['Symbol']:8s}: "
                  f"Holding Cost=${row['Avg Holding Cost']:.2f}/day "
                  f"(Long: ${row['Total Cost Long ($/lot/day)']:+.2f}, "
                  f"Short: ${row['Total Cost Short ($/lot/day)']:+.2f})")

        # 3. LOWEST ENTRY COST (for scalping/day trading)
        print("\n\n3. LOWEST ENTRY COST PAIRS (Best for Scalping/Day Trading):")
        print("-" * 90)

        low_entry_cost = valid_df.sort_values('Total Entry Cost ($/lot)').head(10)

        print("\nPairs with lowest entry/exit costs:")
        for i, (_, row) in enumerate(low_entry_cost.iterrows(), 1):
            print(f"{i:2d}. {row['Symbol']:8s}: "
                  f"Entry Cost=${row['Total Entry Cost ($/lot)']:.2f} "
                  f"(Spread: ${row['Spread Cost ($/lot)']:.2f}, "
                  f"Comm: ${row['Commission ($/lot)']:.2f})")

        # 4. COMMISSION ANALYSIS
        print("\n\n4. COMMISSION ANALYSIS:")
        print("-" * 90)

        # Commission-free pairs
        zero_commission = valid_df[valid_df['Commission ($/lot)'] == 0]
        if not zero_commission.empty:
            print(f"\nCommission-free pairs ({len(zero_commission)}):")
            symbols = list(zero_commission['Symbol'])
            for i in range(0, len(symbols), 10):
                print("  " + ", ".join(symbols[i:i + 10]))

        # High commission pairs
        high_commission = valid_df.sort_values('Commission ($/lot)', ascending=False).head(5)
        print("\nHighest commission per lot:")
        for i, (_, row) in enumerate(high_commission.iterrows(), 1):
            print(f"{i:2d}. {row['Symbol']:8s}: ${row['Commission ($/lot)']:.2f} "
                  f"({row['Commission Type']})")

        # 5. SPREAD ANALYSIS
        print("\n\n5. SPREAD ANALYSIS:")
        print("-" * 90)

        # Tightest spreads
        tight_spreads = valid_df.sort_values('Spread (pips)').head(10)
        print("\nTightest spreads:")
        for i, (_, row) in enumerate(tight_spreads.iterrows(), 1):
            print(f"{i:2d}. {row['Symbol']:8s}: {row['Spread (pips)']} pips "
                  f"(Cost: ${row['Spread Cost ($/lot)']:.2f})")

        # Widest spreads
        wide_spreads = valid_df.sort_values('Spread (pips)', ascending=False).head(10)
        print("\nWidest spreads (avoid for frequent trading):")
        for i, (_, row) in enumerate(wide_spreads.iterrows(), 1):
            print(f"{i:2d}. {row['Symbol']:8s}: {row['Spread (pips)']} pips "
                  f"(Cost: ${row['Spread Cost ($/lot)']:.2f})")

        # 6. TRIPLE SWAP WARNING
        triple_swap = df[df['Triple Swap'] == 'Yes']
        if not triple_swap.empty:
            print("\n\n6. TRIPLE SWAP WARNING:")
            print("-" * 90)
            print(f"These {len(triple_swap)} pairs have TRIPLE swap on rollover days:")
            symbols = list(triple_swap['Symbol'])
            for i in range(0, len(symbols), 10):
                print("  " + ", ".join(symbols[i:i + 10]))

        # 7. SUMMARY STATISTICS
        print("\n\n7. SUMMARY STATISTICS:")
        print("-" * 90)

        total_pairs = len(valid_df)
        avg_spread = valid_df['Spread (pips)'].mean()
        avg_commission = valid_df['Commission ($/lot)'].mean()
        avg_swap_long = valid_df['Swap Long ($/lot/day)'].mean()
        avg_swap_short = valid_df['Swap Short ($/lot/day)'].mean()

        print(f"Total pairs analyzed: {total_pairs}")
        print(f"Average spread: {avg_spread:.1f} pips")
        print(f"Average commission: ${avg_commission:.2f} per lot")
        print(f"Average swap: Long=${avg_swap_long:+.2f}, Short=${avg_swap_short:+.2f} per day")

        # Cost breakdown
        commission_pairs = len(valid_df[valid_df['Commission ($/lot)'] > 0])
        commission_free = len(valid_df[valid_df['Commission ($/lot)'] == 0])
        print(f"Commission: {commission_pairs} pairs have commission, {commission_free} are commission-free")

        print("\n" + "=" * 90)

    def save_to_excel(self, df, filename=None):
        """Save trading cost data to Excel with multiple sheets"""
        if df is None or df.empty:
            print("No data to save")
            return

        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"trading_costs_{timestamp}.xlsx"

        # Create Excel writer
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Save all data
            df.to_excel(writer, sheet_name='All Data', index=False)

            # Save analysis sheets
            for sheet_name, filter_condition in [
                ('Majors', df['Symbol'].isin(['EURUSD', 'GBPUSD', 'USDJPY', 'AUDUSD', 'USDCAD', 'USDCHF', 'NZDUSD'])),
                ('Best_Carry_Trades',
                 (df['Total Cost Long ($/lot/day)'] > 1) | (df['Total Cost Short ($/lot/day)'] > 1)),
                ('Lowest_Holding_Cost',
                 (df['Total Cost Long ($/lot/day)'].abs() < 0.5) &
                 (df['Total Cost Short ($/lot/day)'].abs() < 0.5)),
                ('Lowest_Entry_Cost',
                 df['Total Entry Cost ($/lot)'] < 5),
                ('Commission_Free',
                 df['Commission ($/lot)'] == 0),
                ('Tight_Spreads',
                 df['Spread (pips)'] < 2),
                ('Triple_Swap',
                 df['Triple Swap'] == 'Yes'),
                ('High_Cost_Pairs',
                 (df['Total Entry Cost ($/lot)'] > 20) |
                 (df['Total Cost Long ($/lot/day)'].abs() > 5) |
                 (df['Total Cost Short ($/lot/day)'].abs() > 5))
            ]:
                sheet_df = df[filter_condition]
                if not sheet_df.empty:
                    # Truncate sheet name to 31 characters (Excel limit)
                    sheet_name_short = sheet_name[:31]
                    sheet_df.to_excel(writer, sheet_name=sheet_name_short, index=False)

        print(f"\nData saved to Excel file: {filename}")
        print(f"Sheets created: All Data, Majors, Best_Carry_Trades, Lowest_Holding_Cost,")
        print(f"                Lowest_Entry_Cost, Commission_Free, Tight_Spreads,")
        print(f"                Triple_Swap, High_Cost_Pairs")
        return filename

    def disconnect(self):
        """Disconnect from MT5"""
        mt5.shutdown()
        self.connected = False
        print("\nDisconnected from MT5")


def main():
    # Create monitor instance
    monitor = SwapSpreadCommissionMonitor(MT5_PATH, ACCOUNT, PASSWORD, SERVER)

    try:
        # Connect to MT5
        print("Connecting to MT5...")
        if not monitor.connect():
            print("Failed to connect to MT5. Exiting.")
            return

        # Get all trading costs
        print("\nFetching trading costs for all symbols...")
        cost_df = monitor.get_all_trading_costs()

        if cost_df is not None and not cost_df.empty:
            # Display results
            print("\nGenerating summary report...")
            monitor.display_summary(cost_df)

            # Analyze data
            print("\nPerforming comprehensive analysis...")
            monitor.analyze_trading_costs(cost_df)

            # Save to Excel
            print("\nSaving data to Excel...")
            monitor.save_to_excel(cost_df)

        else:
            print("No trading cost data retrieved")

    except KeyboardInterrupt:
        print("\n\nScript interrupted by user")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Always disconnect
        print("\nCleaning up...")
        monitor.disconnect()


if __name__ == "__main__":
    main()
