export default function validate_(params: Record<string, unknown>): void {
  Object.keys(params).forEach((name) => {
    const value = params[name];
    if (value === undefined || value === null || value === "") {
      throw new Error(`${name} is required.`);
    }
  });
}
