interface RequestUrl {
    protocol: string;
    port: number;
    host: string;
    path: string;
}
interface Request {
    url: RequestUrl;
    headers: any;
    method: string;
}
interface Response {
    statusCode: number;
    body: any;
    headers: any;
    request: Request;
}
interface ErrorResponse {
    response: Response;
    body: any;
}
export declare class ApiError {
    statusCode: number;
    body: any;
    headers: any;
    request: Request;
    constructor(axiosError: any);
    generateError(): ErrorResponse;
}
export {};
