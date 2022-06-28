package ews

import (
	"bytes"
	"crypto/tls"
	"errors"
	"fmt"
	"net/http"
	"regexp"
	"strings"
	"time"

	httpNtlm "github.com/vadimi/go-http-ntlm"
)

var (
	UserName           string // mail or domain\account format
	AccessToken        string
	ExchangeServerAddr string
)

var soapHeader = `<?xml version="1.0" encoding="utf-8" ?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2007_SP1" />
  </soap:Header>
  <soap:Body>
`

func SendMail(to []string, cc []string, topic string, content string) (*http.Response, error) {
	// check username format
	b, err := BuildTextEmail(UserName, to, cc, topic, []byte(content))
	if err != nil {
		return nil, errors.New(fmt.Sprintf("build text email failed: %s", err))
	}

	return Issue(ExchangeServerAddr, UserName, b)
}

func Issue(ewsAddr string, userName string, body []byte) (*http.Response, error) {
	if userName == "" {
		return nil, errors.New("empty user name, please provide valid email or format with domain\\account")
	}

	if ewsAddr == "" {
		return nil, errors.New("empty ews address, please provide valid server address")
	}

	if AccessToken == "" {
		return nil, errors.New("empty ews access token, please provide valid access token")
	}

	b2 := []byte(soapHeader)
	b2 = append(b2, body...)
	b2 = append(b2, "\n  </soap:Body>\n</soap:Envelope>"...)
	req, err := http.NewRequest("POST", ewsAddr, bytes.NewReader(b2))
	if err != nil {
		return nil, errors.New(fmt.Sprintf("create request failed: %s", err))
	}

	var client *http.Client
	re := regexp.MustCompile("^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$")
	isMail := re.MatchString(userName)
	bearer := "Bearer " + AccessToken
	if !isMail {
		// use domain
		l := strings.Split(userName, "\\")
		if len(l) < 2 {
			return nil, errors.New("wrong format of username, not email or format with domain\\account")
		}

		domain := l[0]
		account := l[1]
		client = &http.Client{
			Transport: &httpNtlm.NtlmTransport{
				Domain:          domain,
				User:            account,
				TLSClientConfig: &tls.Config{InsecureSkipVerify: true},
			},
		}
	} else {
		client = &http.Client{
			Transport: &http.Transport{TLSClientConfig: &tls.Config{InsecureSkipVerify: true}},
		}
	}
	req.Header.Add("Authorization", bearer)
	req.Header.Set("Content-Type", "text/xml")
	second := time.Second
	client.Timeout = 10 * second
	client.CheckRedirect = func(req *http.Request, via []*http.Request) error { return http.ErrUseLastResponse }
	return client.Do(req)
}
