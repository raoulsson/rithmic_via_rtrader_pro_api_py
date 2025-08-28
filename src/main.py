"""
R | Trader Pro Local Plugin Implementation
==========================================
Connects locally to R | Trader Pro to access Level 2 DOM and execute orders
WITHOUT expensive API fees ($100+/month)

This implementation uses the plugin architecture where R | Trader Pro acts as
the gateway, and your app connects locally on the same machine.
"""

import socket
import json
import struct
import asyncio
import websocket
import threading
from datetime import datetime
from typing import Dict, List, Optional, Callable
from dataclasses import dataclass
from enum import Enum
import pandas as pd
import xml.etree.ElementTree as ET

# =====================================================
# CONFIGURATION
# =====================================================

@dataclass
class RTraderProConfig:
    """Configuration for local R | Trader Pro connection"""
    # Local connection - no remote API fees!
    host: str = "127.0.0.1"  # Local connection to R | Trader Pro
    port: int = 8000  # Default R | Trader Pro plugin port

    # These should match your R | Trader Pro login
    username: str = ""  # Your AMP/Rithmic username
    gateway: str = "Rithmic 01"  # For live, or "Rithmic Paper Trading" for demo

    # Enable these in R | Trader Pro
    enable_orders: bool = True
    enable_market_data: bool = True
    allow_plugins: bool = True  # MUST BE ENABLED!

# =====================================================
# STEP 1: ENABLE PLUGIN MODE IN R | TRADER PRO
# =====================================================

"""
IMPORTANT SETUP STEPS:
1. Open R | Trader Pro
2. When logging in, make sure:
   - Orders: ON
   - Market Data: ON  
   - Allow Plug-ins: ON (THIS IS CRITICAL!)
3. Login with your AMP credentials
4. Keep R | Trader Pro running while using this plugin
"""

# =====================================================
# LOCAL WEBSOCKET CONNECTION
# =====================================================

class RTraderProConnector:
    """Connects locally to R | Trader Pro via WebSocket"""

    def __init__(self, config: RTraderProConfig):
        self.config = config
        self.ws = None
        self.connected = False
        self.callbacks = {}

    def connect(self):
        """Connect to local R | Trader Pro instance"""
        # Local WebSocket connection - no internet fees!
        ws_url = f"ws://{self.config.host}:{self.config.port}/rithmic"

        try:
            self.ws = websocket.create_connection(ws_url)
            self.connected = True
            print(f"✓ Connected to R | Trader Pro locally at {ws_url}")

            # Start listening thread
            self.listener_thread = threading.Thread(target=self._listen)
            self.listener_thread.daemon = True
            self.listener_thread.start()

            # Send initial handshake
            self._handshake()

        except Exception as e:
            print(f"✗ Failed to connect. Is R | Trader Pro running with plugins enabled?")
            print(f"Error: {e}")
            raise

    def _handshake(self):
        """Initial handshake with R | Trader Pro"""
        handshake_msg = {
            "type": "PLUGIN_CONNECT",
            "plugin_name": "Custom MNQ Trader",
            "version": "1.0",
            "capabilities": ["MARKET_DATA", "ORDER_ENTRY", "LEVEL_2"]
        }
        self.send_message(handshake_msg)

    def send_message(self, message: dict):
        """Send message to R | Trader Pro"""
        if self.ws and self.connected:
            self.ws.send(json.dumps(message))

    def _listen(self):
        """Listen for messages from R | Trader Pro"""
        while self.connected:
            try:
                message = self.ws.recv()
                if message:
                    data = json.loads(message)
                    self._handle_message(data)
            except Exception as e:
                print(f"Listen error: {e}")
                self.connected = False

    def _handle_message(self, data: dict):
        """Route incoming messages to appropriate handlers"""
        msg_type = data.get('type')

        if msg_type in self.callbacks:
            for callback in self.callbacks[msg_type]:
                callback(data)

# =====================================================
# LEVEL 2 DOM HANDLER
# =====================================================

class Level2DOMManager:
    """Manages Level 2 Depth of Market data from R | Trader Pro"""

    def __init__(self, connector: RTraderProConnector):
        self.connector = connector
        self.dom_data = {}
        self.symbol_subscriptions = set()

        # Register callbacks
        connector.callbacks.setdefault('MARKET_DEPTH', []).append(self._on_market_depth)
        connector.callbacks.setdefault('BEST_BID_ASK', []).append(self._on_bba)

    def subscribe_level2(self, symbol: str, exchange: str = "CME"):
        """Subscribe to Level 2 DOM for a symbol"""
        # MNQ is on CME
        subscription_msg = {
            "type": "SUBSCRIBE_MARKET_DEPTH",
            "symbol": symbol,
            "exchange": exchange,
            "depth_levels": 50  # Request 50 levels of depth
        }

        self.connector.send_message(subscription_msg)
        self.symbol_subscriptions.add(f"{symbol}@{exchange}")
        print(f"✓ Subscribed to Level 2 for {symbol}@{exchange}")

    def _on_market_depth(self, data: dict):
        """Handle market depth updates"""
        symbol = data.get('symbol')

        if symbol not in self.dom_data:
            self.dom_data[symbol] = {
                'bids': [],
                'asks': [],
                'last_update': None
            }

        # Update DOM
        dom = self.dom_data[symbol]

        if 'bids' in data:
            dom['bids'] = [(price, size) for price, size in data['bids']]

        if 'asks' in data:
            dom['asks'] = [(price, size) for price, size in data['asks']]

        dom['last_update'] = datetime.now()

        # Print top 5 levels
        self._print_dom_snapshot(symbol)

    def _on_bba(self, data: dict):
        """Handle best bid/ask updates"""
        symbol = data.get('symbol')
        bid = data.get('bid_price')
        ask = data.get('ask_price')
        bid_size = data.get('bid_size')
        ask_size = data.get('ask_size')

        print(f"{symbol}: Bid {bid} ({bid_size}) | Ask {ask} ({ask_size})")

    def _print_dom_snapshot(self, symbol: str):
        """Print current DOM snapshot"""
        if symbol not in self.dom_data:
            return

        dom = self.dom_data[symbol]
        print(f"\n=== {symbol} DOM Snapshot ===")
        print("ASKS (top 5):")
        for price, size in reversed(dom['asks'][:5]):
            print(f"  {price:8.2f} | {size:6d}")
        print("-" * 20)
        print("BIDS (top 5):")
        for price, size in dom['bids'][:5]:
            print(f"  {price:8.2f} | {size:6d}")
        print()

    def get_dom(self, symbol: str) -> dict:
        """Get current DOM for symbol"""
        return self.dom_data.get(symbol, {})

# =====================================================
# ORDER EXECUTION MANAGER
# =====================================================

class OrderManager:
    """Manages order execution through R | Trader Pro"""

    def __init__(self, connector: RTraderProConnector):
        self.connector = connector
        self.orders = {}
        self.positions = {}

        # Register callbacks
        connector.callbacks.setdefault('ORDER_STATUS', []).append(self._on_order_status)
        connector.callbacks.setdefault('FILL', []).append(self._on_fill)
        connector.callbacks.setdefault('POSITION', []).append(self._on_position)

    def place_market_order(self, symbol: str, quantity: int, side: str):
        """Place market order"""
        order_id = f"ORD_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        order_msg = {
            "type": "PLACE_ORDER",
            "order_id": order_id,
            "symbol": symbol,
            "exchange": "CME",
            "quantity": quantity,
            "side": side.upper(),  # BUY or SELL
            "order_type": "MARKET",
            "tif": "DAY"  # Time in force
        }

        self.connector.send_message(order_msg)
        self.orders[order_id] = order_msg
        print(f"✓ Placed market order: {order_id}")
        return order_id

    def place_limit_order(self, symbol: str, quantity: int, side: str, price: float):
        """Place limit order"""
        order_id = f"ORD_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        order_msg = {
            "type": "PLACE_ORDER",
            "order_id": order_id,
            "symbol": symbol,
            "exchange": "CME",
            "quantity": quantity,
            "side": side.upper(),
            "order_type": "LIMIT",
            "limit_price": price,
            "tif": "DAY"
        }

        self.connector.send_message(order_msg)
        self.orders[order_id] = order_msg
        print(f"✓ Placed limit order: {order_id} @ {price}")
        return order_id

    def place_bracket_order(self, symbol: str, quantity: int, side: str,
                            entry_price: float, stop_loss: float, take_profit: float):
        """Place bracket order (entry + stop loss + take profit)"""
        order_id = f"BRK_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        bracket_msg = {
            "type": "PLACE_BRACKET",
            "order_id": order_id,
            "symbol": symbol,
            "exchange": "CME",
            "quantity": quantity,
            "side": side.upper(),
            "entry_type": "LIMIT",
            "entry_price": entry_price,
            "stop_loss_price": stop_loss,
            "take_profit_price": take_profit
        }

        self.connector.send_message(bracket_msg)
        self.orders[order_id] = bracket_msg
        print(f"✓ Placed bracket order: {order_id}")
        print(f"  Entry: {entry_price}, SL: {stop_loss}, TP: {take_profit}")
        return order_id

    def cancel_order(self, order_id: str):
        """Cancel an order"""
        cancel_msg = {
            "type": "CANCEL_ORDER",
            "order_id": order_id
        }
        self.connector.send_message(cancel_msg)
        print(f"✓ Cancellation requested: {order_id}")

    def _on_order_status(self, data: dict):
        """Handle order status updates"""
        order_id = data.get('order_id')
        status = data.get('status')
        print(f"Order {order_id}: {status}")

        if order_id in self.orders:
            self.orders[order_id]['status'] = status

    def _on_fill(self, data: dict):
        """Handle fill notifications"""
        order_id = data.get('order_id')
        fill_price = data.get('fill_price')
        fill_qty = data.get('fill_quantity')
        print(f"⚡ FILL: {order_id} - {fill_qty} @ {fill_price}")

    def _on_position(self, data: dict):
        """Handle position updates"""
        symbol = data.get('symbol')
        position = data.get('position')
        avg_price = data.get('average_price')

        self.positions[symbol] = {
            'quantity': position,
            'avg_price': avg_price,
            'pnl': data.get('unrealized_pnl', 0)
        }

        print(f"Position: {symbol} = {position} @ {avg_price}")

# =====================================================
# MAIN TRADING APPLICATION
# =====================================================

class MNQTradingSystem:
    """Main trading system for MNQ using R | Trader Pro plugin"""

    def __init__(self):
        # Configuration
        self.config = RTraderProConfig(
            username="your_amp_username",  # Update this
            gateway="Rithmic Paper Trading"  # Or "Rithmic 01" for live
        )

        # Initialize components
        self.connector = RTraderProConnector(self.config)
        self.dom_manager = None
        self.order_manager = None

    def start(self):
        """Start the trading system"""
        print("=" * 50)
        print("MNQ Trading System - R | Trader Pro Plugin")
        print("=" * 50)

        # Check R | Trader Pro
        print("\n📋 Pre-flight checklist:")
        print("1. Is R | Trader Pro running? ✓")
        print("2. Is 'Allow Plug-ins' enabled? ✓")
        print("3. Are you logged in with market data? ✓")

        input("\nPress Enter when ready...")

        # Connect
        print("\nConnecting to R | Trader Pro...")
        self.connector.connect()

        # Initialize managers
        self.dom_manager = Level2DOMManager(self.connector)
        self.order_manager = OrderManager(self.connector)

        # Subscribe to MNQ
        self.dom_manager.subscribe_level2("MNQZ24", "CME")  # Dec 2024 contract

        print("\n✓ System ready!")
        print("=" * 50)

    def execute_sample_strategy(self):
        """Example strategy using DOM data"""
        # Get DOM data
        dom = self.dom_manager.get_dom("MNQZ24")

        if not dom or not dom.get('bids') or not dom.get('asks'):
            print("Waiting for DOM data...")
            return

        # Get best bid/ask
        best_bid = dom['bids'][0][0] if dom['bids'] else None
        best_ask = dom['asks'][0][0] if dom['asks'] else None

        if best_bid and best_ask:
            spread = best_ask - best_bid
            mid_price = (best_bid + best_ask) / 2

            print(f"Spread: {spread:.2f} | Mid: {mid_price:.2f}")

            # Example: Place a limit order below the bid
            entry_price = best_bid - 2  # 2 ticks below bid
            self.order_manager.place_limit_order(
                symbol="MNQZ24",
                quantity=1,
                side="BUY",
                price=entry_price
            )

# =====================================================
# USAGE EXAMPLE
# =====================================================

if __name__ == "__main__":
    # Create trading system
    trading_system = MNQTradingSystem()

    # Start the system
    trading_system.start()

    # Run sample strategy
    import time
    try:
        while True:
            trading_system.execute_sample_strategy()
            time.sleep(5)  # Run every 5 seconds
    except KeyboardInterrupt:
        print("\n✓ System stopped")