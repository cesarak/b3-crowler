var Crawler = require("crawler");
const wget = require('wget-improved');
var ProgressBar = require('progress');
const fs = require('fs');

const domain = "http://www.b3.com.br";
const fileName = "partdir_NOVOv2.xls";
const tempDirectory = './tmp/';

console.log("Inicializando");

fs.exists(tempDirectory, exists=>{
    if (!exists) {
        fs.mkdir(tempDirectory,err=>{
        });
    }
})
// let date = new Date();
// let auxDate = date.toLocaleDateString()
//     .replace(/-/g, '');
// let auxTime = date.toLocaleTimeString()
//     .replace(/:/g, '');
// const dateString = auxDate + auxTime
var bar = null;
var c = new Crawler({
    maxConnections: 10,
    // This will be called for each crawled page
    callback: function (error, res, done) {
        if (error) {
            console.log(error);
        } else {
            var $ = res.$;
            // $ is Cheerio by default
            //a lean implementation of core jQuery designed specifically for the server
            console.log($("title").text());
        }
        done();
    }
});
console.log("Acessando página B3");
// Queue URLs with custom callbacks & parameters
c.queue([{
    uri: 'http://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/consultas/mercado-a-vista/participacao-dos-investidores/volume-total/',
    jQuery: false,

    // The global callback won't be called
    callback: function (error, res, done) {
        if (error) {
            console.log("Erro ao acessar site da B3. sys_err: ", error);
        } else {
            // console.log('Grabbed', res.body.length, 'bytes');
            console.log("Recuperando URL do arquivo");
            let endIndex = res.body.indexOf(fileName);
            if (endIndex < 0) {
                console.log("Arquivo não encontrado na página.");
                done();
            }
            let auxPath = res.body.substring(endIndex - 150, endIndex);
            let auxStartIndex = auxPath.indexOf('/data/');
            if (auxStartIndex < 0) {
                console.log("Não encontrada URL do arquivo.");
                done();
            }
            let auxFinalPath = auxPath.substring(auxStartIndex, endIndex);
            let finalPath = domain + auxFinalPath + fileName;
            // console.log(finalPath);

            const output = tempDirectory + fileName;
            const options = {
                // see options below
            };

            console.log("Iniciando download da planilha...");
            let download = wget.download(finalPath, output, options);
            download.on('error', function (err) {
                console.log("Erro ao fazer o download. sys_err:", err);
            });
            download.on('start', function (fileSize) {
                bar = new ProgressBar(':bar', {
                    total: 100
                });
            });
            download.on('end', function (output) {
                // console.log(output);
                console.log("Download da planilha finalizado.");

            });
            download.on('progress', function (progress) {
                typeof progress === 'number'
                bar.tick(progress * 100);
            });
        }
        done();
    }
}]);

// http://www.b3.com.br/data/files/67/51/48/C5/D2E596101F9D6396AC094EA8/partdir_NOVOv2.xls