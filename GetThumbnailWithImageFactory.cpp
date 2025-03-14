HRESULT GetThumbnailWithImageFactory(
    const std::wstring& filePath,
    SIZE size,
    std::vector<byte>& outputBuffer,
    const std::string& output)
{
    auto start = std::chrono::high_resolution_clock::now();
    ComPtr<IShellItem> pShellItem;
    HRESULT hr = SHCreateItemFromParsingName(
      filePath.c_str(), nullptr, IID_PPV_ARGS(&pShellItem));
    if (FAILED(hr))
    {
        std::cout << "FAILED SHCreateItemFromParsingName" << hr << "\n";
        return hr;
    }

    GetShellItemFromIdTime += std::chrono::duration_cast<std::chrono::nanoseconds>(std::chrono::high_resolution_clock::now() - start).count();

    auto startGetThumb = std::chrono::high_resolution_clock::now();
    // Bind to IShellItemImageFactory
    ComPtr<IShellItemImageFactory> pImageFactory;
    hr = pShellItem->QueryInterface(IID_PPV_ARGS(&pImageFactory));
    if (FAILED(hr)) {
        std::wcout << L"Failed to bind to IShellItemImageFactory\n";
        return hr;
    }

    // Get the thumbnail as an HBITMAP
    HBITMAP hBitmap = NULL;
    hr = pImageFactory->GetImage(size, 0, &hBitmap);

    GetThumbnailTime += std::chrono::duration_cast<std::chrono::nanoseconds>(std::chrono::high_resolution_clock::now() - startGetThumb).count();

    if (SUCCEEDED(hr) && hBitmap) {
        hr = HBITMAPToJpeg(hBitmap, outputBuffer);
        if (FAILED(hr) || outputBuffer.empty())
        {
            std::cout << "Error with bitmap " << hr << "\n";
        }
        else if (!output.empty())
        {
            WriteBinaryFile(output, outputBuffer);
        }

        DeleteObject(hBitmap);  // Cleanup after use
    }
    else {
        std::wcout << L"Failed to get thumbnail " << hr;
    }
    return hr;
}
