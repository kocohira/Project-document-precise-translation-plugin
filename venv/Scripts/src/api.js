import axios from 'axios';

const API_URL = 'http://localhost:5000';  // 替换为你的后端URL

export const getTonData = async () => {
  try {
    const response = await axios.get(`${API_URL}/api/ton-data`);
    return response.data;
  } catch (error) {
    console.error('Error fetching TON data:', error);
    throw error;
  }
};