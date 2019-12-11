class TimeoutHandler {
    private static timeoutList = {};

    public static setTimeout(namespace: string, callback: Function, interval: number): void {
        if (this.timeoutList[namespace]) {
            clearTimeout(this.timeoutList[namespace]);
        }

        this.timeoutList[namespace] = setTimeout(callback, interval);
    }

    public static removeTimeout(namespace: string) {
        if (this.timeoutList[namespace]) {
            clearTimeout(this.timeoutList[namespace]);
        }
    }
}

export default TimeoutHandler;