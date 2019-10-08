"Battle sample", Philipp Lenssen
// Being a smaller sample of QL

start
    --- info ? [has gold]
        There's a big monster in front of you
        blocking the path.
        *This could be dangerous!*

    ( monkey.gif )

    % player skill = 10
    % player stamina = 10

    % enemy skill = 10
    % enemy stamina = 10

    $ after win = new path

    --> battle
        Let's start the battle
        Hip hip hooray
        _ player started battle2

battle
    -// This battle functionality can be reused

    % player power = [player skill] + {random 2, 12}
    % enemy power = [enemy skill] + {random 2, 12}

    ? [player power] = [enemy power]
        ---
            Both of your swords miss the other.
            Stamina stays the same.
        --> battle
            Battle again
    ? [player power] > [enemy power]
        ---
            You hit the enemy and wound him.
        % enemy stamina = [enemy stamina] - 2
        --> battle result
            Continue
    ...
        ---
            The enemy hits you and it hurts.
        % player stamina = [player stamina] - 2
        --> battle result
            Continue

battle result
    ? [player stamina] < 1
        ---
            The enemy wounds you deadly.
            **Game over.**
    ? [enemy stamina] < 1
        ---
            You kill the enemy.
            You can continue your game.
        --> [after win]
            You continue
    ...
        ---
            Both of you are still standing.
            Your Stamina is [player stamina], the
            enemy stamina is [enemy stamina].
        --> battle
            Battle again

new path
    ---
        You're on a new path, leaving the monster behind
        you... ||
        *This is just the beginning of the quest,
        but this sample is over.*

information
    +++
        start
        battle*
    ---
        Your stamina is [player stamina]
