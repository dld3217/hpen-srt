// "Good afternoon, Dennis" — same greeting as POC Manager (handles "Last, First" display names).
export const userGreeting = (displayName: string): string => {
  const h = new Date().getHours();
  const tod = h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening';
  const n = (displayName || '').includes(',') ? displayName.split(',')[1].trim().split(' ')[0] : (displayName || '').split(' ')[0];
  return n ? `${tod}, ${n}` : tod;
};
