"Simple sample", Philipp Lenssen

start
    ---
        There's a treasure chest here, and
        a door leads to the north.

    --> open chest ? ![chest open]
        Open the treasure chest

    --> northern room
        Leave the room

open chest
    ---
        You open the chest and take two
        silver coins.

    _ chest open
    % silver = [silver] + 2
    $ test = hello world

    <--
        Continue

northern room
    ---
        The path leads north.
        You walk a bit, get tired, go to
        sleep & dream about your [silver] silver coins.
        *The End*