import React, { useState, useEffect } from 'react';
import { TonConnectUI } from '@tonconnect/ui';
import { getTonData } from './api';

const tonConnectUI = new TonConnectUI({
    manifestUrl: 'http://localhost:3000/tonconnect-manifest.json'
  });

function TonWallet() {
  const [wallet, setWallet] = useState(null);
  const [tonData, setTonData] = useState(null);

  useEffect(() => {
    tonConnectUI.onStatusChange(setWallet);
  }, []);

  const handleConnect = () => {
    tonConnectUI.connectWallet();
  };

  const fetchTonData = async () => {
    try {
      const data = await getTonData();
      setTonData(data);
    } catch (error) {
      console.error('Error:', error);
    }
  };

  return (
    <div>
      {wallet ? (
        <div>
          <p>Connected: {wallet.account.address}</p>
          <button onClick={fetchTonData}>Fetch TON Data</button>
          {tonData && <p>TON Data: {JSON.stringify(tonData)}</p>}
        </div>
      ) : (
        <button onClick={handleConnect}>Connect Wallet</button>
      )}
    </div>
  );
}

export default TonWallet;

