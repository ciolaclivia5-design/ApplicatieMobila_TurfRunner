import { Platform } from 'react-native';

const API_HOST = Platform.OS === 'ios' ? 'localhost' : '10.0.2.2';

export const API_BASE_URL = `http://${API_HOST}:3000`;
