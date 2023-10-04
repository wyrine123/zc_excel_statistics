.PHONY: build clean tool lint help

BINARY_NAME=zc-excel-statistics

build:
	go mod tidy
	GOARCH=amd64 GOOS=darwin go build -o build/package/${BINARY_NAME}-darwin main.go
	GOARCH=amd64 GOOS=linux go build -o build/package/${BINARY_NAME}-linux main.go
	GOARCH=amd64 GOOS=windows go build -o build/package/${BINARY_NAME}-windows main.go
	go build -o build/package/${BINARY_NAME} main.go

run: build
	./build/package/${BINARY_NAME}

clean:
	go clean
	rm build/package/${BINARY_NAME}-darwin
	rm build/package/${BINARY_NAME}-linux
	rm build/package/${BINARY_NAME}-windows
	rm build/package/${BINARY_NAME}

lint:
	golangci-lint run