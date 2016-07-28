import {Injectable} from '@angular/core';
import {Http} from '@angular/http';
import {Utilities, RequestHelper} from '../helpers';

export class Snippet {
    meta: {
        name: string,
        id: string
    };
    ts: string;
    html: string;
    css: string;
    extras: string;

    hash: string;
    jsHash: string;

    private _compiledJs: string;

    constructor(meta, ts, html, css, extras) {
        this.meta = meta;
        this.ts = ts;
        this.css = css;
        this.extras = extras;
        this.html = html;
    }

    // A bit of a hack (probably doesn't belong here, but want to get an easy "run" link)
    get runUrl(): string {
        var url = window.location.toString() + "#/run/" + this.meta.id;
        console.log(url);
        return url;
    }

    get js(): Promise<string> {
        if (Snippet._isPureValidJs(this.ts)) {
            this._compiledJs = this.ts;
            return Promise.resolve(this._compiledJs);        
        }
        else {
            // FIXME expose to user
            console.log(this.ts);
            throw Utilities.error("Invalid JavaScript (or is TypeScript, which we don't have a compiler for yet)")
            // return this._compile(this.ts).then((compiledJs) => {
            //     this._compiledJs = compiledJs;
            //     return compiledJs; 
            // })
        }
    }

    getJsLibaries(): Array<string> {
        // FIXME
        return [
            "https://appsforoffice.microsoft.com/lib/1/hosted/office.js",
            "https://npmcdn.com/jquery",
            "https://npmcdn.com/office-ui-fabric/dist/js/jquery.fabric.min.js",
        ];
    }

    getCssStylesheets(): Array<string> {
        // FIXME
        return [
            "https://npmcdn.com/office-ui-fabric/dist/css/fabric.min.css",
            "https://npmcdn.com/office-ui-fabric/dist/css/fabric.components.min.css",
        ];
    }

    static _isPureValidJs(scriptText): boolean {
        try {
			new Function(scriptText);
            return true;
        } catch (syntaxError) {
            return false;
        }
    }

    private _compile(ts: string): Promise<string> {
        // FIXME
        return Promise.resolve(ts);
    }

    private _hash() {
        // FIXME
    }

    static create(meta, js, html, css, extras): Promise<Snippet> {
        return Promise.all([meta, js, html, css, extras])
            .then(results => new Snippet(results[0], results[1], results[2], results[3], results[4]))
            .catch(error => Utilities.error);
    }    
}

@Injectable()
export class SnippetsService {
    private _baseUrl: string = 'https://xlsnippets.azurewebsites.net/api';

    constructor(private _request: RequestHelper) {

    }

    get(snippetId: string): Promise<Snippet> {
        console.log(snippetId);
        var meta = this._request.get(this._baseUrl + '/snippets/' + snippetId);
        var js = this._request.get(this._baseUrl + '/snippets/' + snippetId + '/content/js', null, true);
        var html = this._request.get(this._baseUrl + '/snippets/' + snippetId + '/content/html', null, true);
        var css = this._request.get(this._baseUrl + '/snippets/' + snippetId + '/content/css', null, true);
        var extras = this._request.get(this._baseUrl + '/snippets/' + snippetId + '/content/extras', null, true);
        return Snippet.create(meta, js, html, css, extras);
    }

    create(name: string, password?: string): Promise<{ id: string, password: string }> {
        var body = { name: name, password: password };

        return this._request.post(this._baseUrl + '/snippets', body)
            .then((data: any) => {
                return {
                    id: data.id,
                    password: data.password
                }
            })
    }

    uploadContent(snippetId: string, password: string, fileName: string, content: string) {
        var headers = RequestHelper.generateHeaders({
            "Content-Type": "application/octet-stream",
            "x-ms-b64-password": btoa(password)
        });
        return this._request.putRaw(this._baseUrl + '/snippets/' + snippetId + '/content/' + fileName, content, headers);
    }
}