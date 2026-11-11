import { PermissionsAndroid, Platform } from 'react-native';

export async function requestLocationPermission() {
  if (Platform.OS !== 'android') return true;

  try {
    const result = await PermissionsAndroid.request(
      PermissionsAndroid.PERMISSIONS.ACCESS_FINE_LOCATION
    );
    return result === PermissionsAndroid.RESULTS.GRANTED;
  } catch (error) {
    console.warn('Failed to request location permission:', error);
    return false;
  }
}
