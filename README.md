# FIPE API usage

[https://apifipe.com.br/doc.php](https://apifipe.com.br/doc.php)

```bash

# api token
export TOKEN=""

# vehicle types => carros, motos, caminhoes

# brands info
# https://apifipe.com.br/api/{vehicleType}/${TOKEN}
# e.g.:
curl -s -X GET --url "https://apifipe.com.br/api/carros/${TOKEN}"


# models info
# https://apifipe.com.br/api/{vehicleType}/{brandCode}/${TOKEN}
# e.g.:
curl -s -X GET --url "https://apifipe.com.br/api/carros/23/${TOKEN}"


# years per model info
# https://apifipe.com.br/api/{vehicleType}/{brandCode}/{modelCode}/${TOKEN}
# e.g.:
curl -s -X GET --url "https://apifipe.com.br/api/carros/23/8825/${TOKEN}"


# fipe info
# https://apifipe.com.br/api/{vehicleType}/{brandCode}/{modelCode}/{yearCode}/${TOKEN}
# e.g.:
curl -s -X GET --url "https://apifipe.com.br/api/carros/23/8825/2021-1/${TOKEN}"


# fipe info by code
# https://apifipe.com.br/api/fipe/{fipeCode}/${TOKEN}
# e.g.:
curl -s -X GET --url "https://apifipe.com.br/api/fipe/004501-2/${TOKEN}" | jq -r '.'

```
