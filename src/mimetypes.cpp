// The file is automatically generated, changes will be overwritten, 
// from the mimetypes library.
//
// Copyright (c) 2023, zero.kwok@foxmail.com
// For the full copyright and license information, please view the LICENSE
// file that was distributed with this source code.
//
// https://github.com/ZeroKwok/mimetypes
//
// Ported by Zbigniew Minciel for inclusion with free Windows MBox Viewer

#include "mimetypes.h"

std::string mimetypes::from_extension(const std::string& extension, bool strict)
{
    if (extension.empty())
        return strict ? "" : "application/octet-stream";

    auto _lower = [](std::string s) -> std::string {
        std::transform(s.begin(), s.end(), s.begin(),
            [](unsigned char c) { return std::tolower(c); }
        );
        return s;
    };

    auto _get = [&](auto name) {
        auto item = detail::__mimetypes.find(name);
        if (item != detail::__mimetypes.end())
            return item->second;
        return strict ? "" : "application/octet-stream";
    };

    auto extension1 = _lower(extension);
    if (extension1[0] == '.' && extension1.size() > 1)
        return _get(extension1.c_str() + 1);
    else
        return _get(extension1.c_str());
}


