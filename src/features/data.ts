import dataProtocol from './protocol.json'
import dataCertificate from './certificate.json'

type Protocol = typeof dataProtocol
type Certificate = typeof dataCertificate
export const getDataProtocol = (): Protocol => {
    return dataProtocol
}

export const getDataCertificate = (): Certificate => {
    return dataCertificate
}